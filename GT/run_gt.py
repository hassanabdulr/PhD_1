#!/usr/bin/env python
"""
Graph Theory Analysis Pipeline for XCP-D outputs.

Usage:
    # Run single parcellation:
    python run_gt.py --parcellation Glasser
    
    # Run multiple parcellations:
    python run_gt.py --parcellation Glasser Gordon 4S456Parcels
    
    # Run all available parcellations:
    python run_gt.py --all
"""

import argparse
import numpy as np
import pandas as pd
import nibabel as nib
import bct  # Brain Connectivity Toolbox
from scipy import stats
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for cluster/batch jobs
import matplotlib.pyplot as plt
import seaborn as sns
from nilearn.connectome import ConnectivityMeasure
import networkx as nx
from pathlib import Path
import os
import re
from tqdm import tqdm


# Set style for visualizations
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("husl")

# === CONFIGURATION ===
XCP_OUTPUT_DIR = '/projects/aabdulrasul/TAY/GT/data'
OUTPUT_DIR = '/projects/aabdulrasul/TAY/GT/test'

EXPECTED_PARCELS = {
    '4S156Parcels': 156, '4S256Parcels': 256, '4S356Parcels': 356,
    '4S456Parcels': 456, '4S556Parcels': 556, '4S656Parcels': 656,
    '4S756Parcels': 756, '4S856Parcels': 856, '4S956Parcels': 956,
    '4S1056Parcels': 1056, 'Glasser': 360, 'Gordon': 333,
    'HCP': 360, 'MIDB': 82, 'MyersLabonte': 246, 'Tian': 54
}

USE_CONCATENATED = False
EDGE_TYPE = 'positive_only'
THRESHOLDS = [0.05, 0.10, 0.15, 0.20, 0.25, 0.30, 0.35, 0.40]
PRIMARY_THRESHOLD = 0.15
THRESHOLD_TYPE = 'proportional'
COMPUTE_SMALLWORLD = True
N_RANDOM = 100

np.random.seed(23)




# Cell 3: XCP-D File Discovery Functions
def find_xcp_subjects(xcp_dir):
    """Find all subjects in XCP-D output directory."""
    xcp_path = Path(xcp_dir)
    subjects = sorted([d.name for d in xcp_path.glob('sub-*') if d.is_dir()])
    return subjects

def find_sessions(subject_dir):
    """Find all sessions for a subject."""
    return sorted([d.name for d in Path(subject_dir).glob('ses-*') if d.is_dir()])

def find_ptseries_files(func_dir, parcellation=None, concatenated=True, task='rest'):
    """
    Find ptseries files in a func directory.
    """
    func_path = Path(func_dir)
    
    pattern = '*_stat-mean_timeseries.ptseries.nii'
    files = list(func_path.glob(pattern))
    
    if task:
        files = [f for f in files if f'task-{task}' in f.name]
    
    if concatenated:
        files = [f for f in files if '_run-' not in f.name]
    else:
        files = [f for f in files if '_run-' in f.name]
    
    if parcellation:
        files = [f for f in files if f'seg-{parcellation}' in f.name]
    
    files = sorted(files)
    
    # Only warn for concatenated mode where multiple files is unexpected
    if concatenated and len(files) > 1:
        print(f"  Warning: Multiple concatenated files found for {parcellation}")
    
    return files

def get_available_parcellations(func_dir, concatenated=True):
    """Get list of all available parcellations in a func directory."""
    files = find_ptseries_files(func_dir, parcellation=None, concatenated=concatenated)
    parcellations = set()
    
    for f in files:
        match = re.search(r'seg-([^_]+)_', f.name)
        if match:
            parcellations.add(match.group(1))
    
    return sorted(parcellations)

def load_ptseries(filepath, parcellation=None, expected_parcels=None):
    """
    Load parcellated time series from XCP-D ptseries file.
    
    CIFTI ptseries convention: (timepoints, parcels)
    
    Parameters:
    -----------
    filepath : str or Path
        Path to ptseries file
    parcellation : str, optional
        Parcellation name for lookup in expected_parcels
    expected_parcels : dict, optional
        Dictionary mapping parcellation names to expected parcel counts
    """
    img = nib.load(filepath)
    data = img.get_fdata()
    
    if data.ndim == 1:
        raise ValueError(f"1D data in {filepath}")
    
    # Robust orientation check
    if expected_parcels and parcellation and parcellation in expected_parcels:
        n_expected = expected_parcels[parcellation]
        if data.shape[1] == n_expected:
            pass  # Correct orientation: (timepoints, parcels)
        elif data.shape[0] == n_expected:
            print(f"  Note: Transposing data from shape {data.shape} (expected {n_expected} parcels)")
            data = data.T
        else:
            print(f"  Warning: Neither dimension matches expected {n_expected} parcels for {parcellation}")
            # Fall back to heuristic
            if data.shape[0] < data.shape[1]:
                data = data.T
    else:
        # Heuristic fallback: assume more timepoints than parcels
        if data.shape[0] < data.shape[1]:
            print(f"  Note: Transposing data from shape {data.shape}")
            data = data.T
    
    return data


# Cell 4: Connectivity Matrix Functions (CORRECTED)
def compute_connectivity_matrix(timeseries, method='pearson', fisher_z=True):
    """
    Compute functional connectivity matrix from parcellated time series.
    Returns Fisher z-transformed values by default (fisher_z=True).
    The z-transform (z = arctanh(r)) normalizes the sampling distribution
    of correlations, which is essential for valid group-level averaging
    and statistical inference across subjects.
    """
    
    if method == 'pearson':
        conn_matrix = np.corrcoef(timeseries.T)
        
        if fisher_z:
            # Avoid infinities at r = ±1
            conn_matrix = np.clip(conn_matrix, -0.999999, 0.999999)
            conn_matrix = np.arctanh(conn_matrix)  # Fisher r-to-z
    
    elif method == 'partial':
        measure = ConnectivityMeasure(kind='partial correlation')
        conn_matrix = measure.fit_transform([timeseries])[0]
        
        if fisher_z:
            conn_matrix = np.clip(conn_matrix, -0.999999, 0.999999)
            conn_matrix = np.arctanh(conn_matrix)

    # Remove diagonal AFTER transform
    np.fill_diagonal(conn_matrix, 0)
    
    # Clean up
    conn_matrix = np.nan_to_num(conn_matrix)
    conn_matrix = (conn_matrix + conn_matrix.T) / 2  # Enforce symmetry
    
    return conn_matrix

def prepare_graph_matrix(conn_matrix, edge_type='positive_only'):
    """
    Prepare connectivity matrix for graph analysis based on edge handling choice.
    """
    if edge_type == 'positive_only':
        graph_matrix = np.clip(conn_matrix, 0, None)
    elif edge_type == 'absolute':
        graph_matrix = np.abs(conn_matrix)
    elif edge_type == 'signed':
        graph_matrix = conn_matrix.copy()
    else:
        raise ValueError(f"Unknown edge_type: {edge_type}")
    
    # Ensure symmetry and zero diagonal
    graph_matrix = (graph_matrix + graph_matrix.T) / 2
    np.fill_diagonal(graph_matrix, 0)
    
    return graph_matrix

def threshold_matrix(conn_matrix, threshold, threshold_type='proportional', binarize=False):
    """
    Threshold connectivity matrix.
    
    Works on the prepared graph matrix (Fisher z-values after edge handling).
    Proportional thresholding is rank-based and thus invariant to monotonic
    transforms, so results are equivalent whether applied to r or z values.
    """
    if threshold_type == 'proportional':
        thresholded = bct.threshold_proportional(conn_matrix, threshold)
    elif threshold_type == 'absolute':
        thresholded = conn_matrix.copy()
        thresholded[thresholded < threshold] = 0
    else:
        raise ValueError(f"Unknown threshold_type: {threshold_type}")
    
    # Ensure symmetry and zero diagonal after thresholding
    thresholded = (thresholded + thresholded.T) / 2
    np.fill_diagonal(thresholded, 0)
    
    if binarize:
        thresholded = (thresholded > 0).astype(float)
    
    return thresholded

def compute_density(A):
    """
    Compute density of an undirected graph correctly.
    
    For undirected graphs: density = m / (N*(N-1)/2)
    where m = number of edges (upper triangle only)
    """
    np.fill_diagonal(A, 0)
    m = np.sum(np.triu(A, 1) > 0)  # Count edges in upper triangle only
    n = A.shape[0]
    max_edges = n * (n - 1) / 2
    return m / max_edges if max_edges > 0 else 0

# Cell 5: Graph Theory Metrics (CORRECTED)
def compute_weighted_metrics(W):
    """
    Compute graph metrics on WEIGHTED network.
    
    Parameters:
    -----------
    W : np.ndarray
        Weighted adjacency matrix (non-negative, thresholded)
        
    Returns:
    --------
    metrics : dict
        Dictionary of weighted graph metrics
    """
    metrics = {}
    
    # Ensure non-negative weights, symmetry, zero diagonal
    W = np.clip(W, 0, None)
    W = (W + W.T) / 2
    np.fill_diagonal(W, 0)
    
    # 1. Global Efficiency (weighted)
    metrics['global_efficiency'] = bct.efficiency_wei(W)
    
    # 2. Clustering Coefficient (weighted)
    clustering = bct.clustering_coef_wu(W)
    metrics['clustering_nodal'] = clustering
    metrics['clustering_mean'] = np.nanmean(clustering)
    
    # 3. Characteristic Path Length using BCT standard conversion
    # Use BCT's weight_conversion for consistency with literature
    D = bct.weight_conversion(W, 'lengths')
    D[W == 0] = np.inf
    np.fill_diagonal(D, 0)
    
    try:
        charpath, efficiency, ecc, radius, diameter = bct.charpath(D, include_infinite=False)
        metrics['characteristic_path_length'] = charpath if np.isfinite(charpath) else np.nan
        metrics['radius'] = radius if np.isfinite(radius) else np.nan
        metrics['diameter'] = diameter if np.isfinite(diameter) else np.nan
    except:
        charpath, efficiency, ecc, radius, diameter = bct.charpath(D)
        metrics['characteristic_path_length'] = charpath if np.isfinite(charpath) else np.nan
        metrics['radius'] = radius if np.isfinite(radius) else np.nan
        metrics['diameter'] = diameter if np.isfinite(diameter) else np.nan
    
    # 4. Modularity - FIXED: best of 10 runs for reproducibility
    best_q = -np.inf
    best_ci = None
    for _ in range(10):
        ci_temp, q_temp = bct.community_louvain(W)
        if q_temp > best_q:
            best_q = q_temp
            best_ci = ci_temp
    ci, q = best_ci, best_q
    metrics['modularity'] = q
    metrics['community_assignments'] = ci
    metrics['n_modules'] = len(np.unique(ci))

    # 5. Node Strength
    strength = bct.strengths_und(W)
    metrics['strength'] = strength
    metrics['strength_mean'] = np.mean(strength)
    
    # 6. Betweenness Centrality (weighted)
    try:
        bc = bct.betweenness_wei(D)
        bc = np.where(np.isinf(bc), 0, bc)
    except:
        bc = np.zeros(W.shape[0])
    metrics['betweenness_centrality'] = bc
    metrics['betweenness_mean'] = np.nanmean(bc)
    
    # 7. Participation Coefficient
    pc = bct.participation_coef(W, ci)
    metrics['participation_coefficient'] = pc
    metrics['participation_mean'] = np.nanmean(pc)
    
    # 8. Local Efficiency (weighted)
    local_eff = bct.efficiency_wei(W, local=True)
    metrics['local_efficiency'] = local_eff
    metrics['local_efficiency_mean'] = np.nanmean(local_eff)
    
    # 9. Assortativity (weighted)
    try:
        metrics['assortativity'] = bct.assortativity_wei(W, flag=0)
    except:
        metrics['assortativity'] = np.nan
    
    # 10. Transitivity (weighted)
    metrics['transitivity'] = bct.transitivity_wu(W)
    
    return metrics

def compute_binary_metrics(A):
    """
    Compute graph metrics on BINARY network.
    """
    metrics = {}
    
    # Ensure binary, symmetric, zero diagonal
    A = (A > 0).astype(float)
    A = (A + A.T) / 2
    A = (A > 0).astype(float)
    np.fill_diagonal(A, 0)
    
    # 1. Global Efficiency
    metrics['global_efficiency'] = bct.efficiency_bin(A)
    
    # 2. Clustering Coefficient
    clustering = bct.clustering_coef_bu(A)
    metrics['clustering_nodal'] = clustering
    metrics['clustering_mean'] = np.nanmean(clustering)
    
    # 3. Characteristic Path Length
    D = bct.distance_bin(A)
    try:
        charpath, efficiency, ecc, radius, diameter = bct.charpath(D, include_infinite=False)
    except:
        charpath, efficiency, ecc, radius, diameter = bct.charpath(D)
    
    metrics['characteristic_path_length'] = charpath if np.isfinite(charpath) else np.nan
    metrics['radius'] = radius if np.isfinite(radius) else np.nan
    metrics['diameter'] = diameter if np.isfinite(diameter) else np.nan
    
    # 4. Modularity - FIXED: Use A, not W; best of 10 runs
    best_q = -np.inf
    best_ci = None
    for _ in range(10):
        ci_temp, q_temp = bct.community_louvain(A)  # FIXED: was W, now A
        if q_temp > best_q:
            best_q = q_temp
            best_ci = ci_temp
    ci, q = best_ci, best_q
    metrics['modularity'] = q
    metrics['community_assignments'] = ci
    metrics['n_modules'] = len(np.unique(ci))
    
    # 5. Degree
    degree = bct.degrees_und(A)
    metrics['degree'] = degree
    metrics['degree_mean'] = np.mean(degree)
    
    # 6. Betweenness Centrality
    bc = bct.betweenness_bin(A)
    metrics['betweenness_centrality'] = bc
    metrics['betweenness_mean'] = np.nanmean(bc)
    
    # 7. Participation Coefficient
    pc = bct.participation_coef(A, ci)
    metrics['participation_coefficient'] = pc
    metrics['participation_mean'] = np.nanmean(pc)
    
    # 8. Local Efficiency
    local_eff = bct.efficiency_bin(A, local=True)
    metrics['local_efficiency'] = local_eff
    metrics['local_efficiency_mean'] = np.nanmean(local_eff)
    
    # 9. Assortativity
    try:
        metrics['assortativity'] = bct.assortativity_bin(A, flag=0)
    except:
        metrics['assortativity'] = np.nan
    
    # 10. Transitivity
    metrics['transitivity'] = bct.transitivity_bu(A)
    
    # 11. Rich Club (only valid for binary)
    try:
        rc = bct.rich_club_bu(A)
        metrics['rich_club'] = rc
    except:
        metrics['rich_club'] = np.array([])
    
    return metrics

def compute_small_worldness_binary(A, n_random=N_RANDOM):
    """
    Compute small-worldness on BINARY network using proper null model.
    
    This is the CORRECT approach: randmio_und is valid for binary networks.
    """
    # Ensure binary, symmetric, zero diagonal
    A = (A > 0).astype(float)
    A = (A + A.T) / 2
    A = (A > 0).astype(float)
    np.fill_diagonal(A, 0)
    
    # Real network metrics
    C_real = np.mean(bct.clustering_coef_bu(A))
    D_real = bct.distance_bin(A)
    
    try:
        L_real, _, _, _, _ = bct.charpath(D_real, include_infinite=False)
    except:
        L_real, _, _, _, _ = bct.charpath(D_real)
    
    if not np.isfinite(L_real):
        print("  Warning: Network disconnected, cannot compute small-worldness")
        return {
            'sigma': np.nan, 'gamma': np.nan, 'lambda': np.nan,
            'C_real': C_real, 'C_rand': np.nan, 'L_real': np.nan, 'L_rand': np.nan
        }
    
    # Random networks (degree-preserving rewiring)
    C_rand_list = []
    L_rand_list = []
    
    for _ in range(n_random):
        try:
            rand_net, _ = bct.randmio_und(A, 10)
            
            C_rand = np.mean(bct.clustering_coef_bu(rand_net))
            D_rand = bct.distance_bin(rand_net)
            
            try:
                L_rand, _, _, _, _ = bct.charpath(D_rand, include_infinite=False)
            except:
                L_rand, _, _, _, _ = bct.charpath(D_rand)
            
            if np.isfinite(L_rand):
                C_rand_list.append(C_rand)
                L_rand_list.append(L_rand)
        except Exception:
            continue
    
    if len(C_rand_list) < 10:
        print(f"  Warning: Only {len(C_rand_list)} valid random networks generated")
        if len(C_rand_list) == 0:
            return {
                'sigma': np.nan, 'gamma': np.nan, 'lambda': np.nan,
                'C_real': C_real, 'C_rand': np.nan, 'L_real': L_real, 'L_rand': np.nan
            }
    
    C_rand_mean = np.mean(C_rand_list)
    L_rand_mean = np.mean(L_rand_list)
    
    # Safety check for division
    if C_rand_mean > 0 and L_rand_mean > 0:
        gamma = C_real / C_rand_mean
        lam = L_real / L_rand_mean
        sigma = gamma / lam
    else:
        gamma = np.nan
        lam = np.nan
        sigma = np.nan
    
    return {
        'sigma': sigma, 'gamma': gamma, 'lambda': lam,
        'C_real': C_real, 'C_rand': C_rand_mean, 'L_real': L_real, 'L_rand': L_rand_mean
    }


# Cell 6: Visualization Functions
def plot_connectivity_matrix(conn_matrix, title="Connectivity Matrix", save_path=None):
    """Plot connectivity matrix as heatmap."""
    fig, ax = plt.subplots(figsize=(10, 8))
    
    vmax = np.percentile(np.abs(conn_matrix), 95)
    im = ax.imshow(conn_matrix, cmap='RdBu_r', vmin=-vmax, vmax=vmax, aspect='equal')
    
    cbar = plt.colorbar(im, ax=ax, shrink=0.8)
    cbar.set_label('Correlation (r)', fontsize=11)
    
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel('Parcels', fontsize=11)
    ax.set_ylabel('Parcels', fontsize=11)
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig

def plot_community_matrix(conn_matrix, community_labels, title="Community Structure", save_path=None):
    """Plot connectivity matrix reordered by community."""
    sorted_idx = np.argsort(community_labels)
    sorted_matrix = conn_matrix[np.ix_(sorted_idx, sorted_idx)]
    sorted_communities = community_labels[sorted_idx]
    
    fig, ax = plt.subplots(figsize=(10, 8))
    
    vmax = np.percentile(np.abs(sorted_matrix), 95)
    im = ax.imshow(sorted_matrix, cmap='RdBu_r', vmin=-vmax, vmax=vmax, aspect='equal')
    
    # Add community boundaries
    unique_communities = np.unique(sorted_communities)
    for comm in unique_communities:
        idx = np.where(sorted_communities == comm)[0]
        if len(idx) > 0:
            boundary = idx[-1] + 0.5
            ax.axhline(boundary, color='black', linewidth=1.5)
            ax.axvline(boundary, color='black', linewidth=1.5)
    
    cbar = plt.colorbar(im, ax=ax, shrink=0.8)
    cbar.set_label('Correlation (r)', fontsize=11)
    
    ax.set_title(title, fontsize=12, fontweight='bold')
    ax.set_xlabel('Parcels (sorted by community)', fontsize=11)
    ax.set_ylabel('Parcels (sorted by community)', fontsize=11)
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig

def plot_metrics_summary(metrics_df, parcellation, save_path=None):
    """Create summary visualization of metrics distribution."""
    metrics_to_plot = ['global_efficiency_w', 'clustering_mean_w', 'modularity_w',
                      'characteristic_path_length_w', 'transitivity_w', 'small_worldness_sigma']
    
    # Fallback to available columns
    available_metrics = [m for m in metrics_to_plot if m in metrics_df.columns]
    
    if len(available_metrics) == 0:
        # Try without suffix
        metrics_to_plot = ['global_efficiency', 'clustering_mean', 'modularity',
                          'characteristic_path_length', 'transitivity', 'small_worldness_sigma']
        available_metrics = [m for m in metrics_to_plot if m in metrics_df.columns]
    
    n_metrics = len(available_metrics)
    
    if n_metrics == 0:
        print("No metrics to plot")
        return None
    
    fig, axes = plt.subplots(2, 3, figsize=(14, 8))
    axes = axes.flatten()
    
    for i, metric in enumerate(available_metrics):
        ax = axes[i]
        # Remove inf and nan values
        data = metrics_df[metric].replace([np.inf, -np.inf], np.nan).dropna()
        
        if len(data) == 0:
            ax.text(0.5, 0.5, f'No valid data', ha='center', va='center', transform=ax.transAxes)
            ax.set_title(metric.replace('_', ' ').title(), fontsize=10)
            continue
        
        ax.hist(data, bins=min(20, max(5, len(data)//2)), alpha=0.7, color='steelblue', edgecolor='black')
        ax.axvline(np.mean(data), color='red', linestyle='--', linewidth=2)
        ax.set_xlabel(metric.replace('_', ' ').title(), fontsize=10)
        ax.set_ylabel('Count', fontsize=10)
        ax.set_title(f'μ = {np.mean(data):.3f} ± {np.std(data):.3f}', fontsize=10)
    
    for j in range(len(available_metrics), len(axes)):
        axes[j].set_visible(False)
    
    plt.suptitle(f'Graph Metrics Distribution - {parcellation}\n(n = {len(metrics_df)} subjects)', 
                 fontsize=12, fontweight='bold')
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig

def plot_threshold_sweep(sweep_df, subject_id, save_path=None):
    """Plot metrics across different thresholds (AUC analysis)."""
    metrics_to_plot = ['global_efficiency_w', 'clustering_mean_w', 'modularity_w',
                      'characteristic_path_length_w', 'transitivity_w', 'small_worldness_sigma']
    
    available = [m for m in metrics_to_plot if m in sweep_df.columns]
    
    fig, axes = plt.subplots(2, 3, figsize=(14, 8))
    axes = axes.flatten()
    
    for i, metric in enumerate(available):
        ax = axes[i]
        data = sweep_df[['threshold', metric]].dropna()
        
        if len(data) > 0:
            ax.plot(data['threshold'], data[metric], 'o-', linewidth=2, markersize=8, color='steelblue')
            
            # Compute AUC
            if len(data) > 1:
                auc = np.trapezoid(data[metric], data['threshold'])
                ax.set_title(f'{metric.replace("_", " ").title()}\nAUC = {auc:.4f}', fontsize=10)
            else:
                ax.set_title(metric.replace('_', ' ').title(), fontsize=10)
        
        ax.set_xlabel('Threshold (density)', fontsize=10)
        ax.set_ylabel('Value', fontsize=10)
        ax.grid(True, alpha=0.3)
    
    for j in range(len(available), len(axes)):
        axes[j].set_visible(False)
    
    plt.suptitle(f'Threshold Sweep - {subject_id}', fontsize=12, fontweight='bold')
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig

def plot_strength_distribution(strength_values, title="Strength Distribution", save_path=None):
    """Plot node strength distribution."""
    fig, axes = plt.subplots(1, 2, figsize=(12, 4))
    
    axes[0].hist(strength_values, bins=30, density=True, alpha=0.7, 
                 color='steelblue', edgecolor='black')
    axes[0].set_xlabel('Node Strength', fontsize=11)
    axes[0].set_ylabel('Density', fontsize=11)
    axes[0].set_title('Histogram', fontsize=11, fontweight='bold')
    
    hist, bins = np.histogram(strength_values, bins=30)
    bin_centers = (bins[:-1] + bins[1:]) / 2
    mask = hist > 0
    if np.any(mask):
        axes[1].loglog(bin_centers[mask], hist[mask], 'o', markersize=6, color='steelblue')
    axes[1].set_xlabel('Node Strength (log)', fontsize=11)
    axes[1].set_ylabel('Frequency (log)', fontsize=11)
    axes[1].set_title('Log-Log Plot', fontsize=11, fontweight='bold')
    
    plt.suptitle(title, fontsize=12, fontweight='bold')
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig

def plot_network_graph(conn_matrix, community_labels=None, node_metric=None,
                       title="Network Graph", save_path=None, top_percent=30):
    """
    Create network visualization using NetworkX.
    
    Parameters:
    -----------
    conn_matrix : np.ndarray
        Connectivity matrix (thresholded, weighted)
    community_labels : np.ndarray, optional
        Community assignments for node coloring
    node_metric : np.ndarray, optional
        Metric for node sizing (e.g., strength, betweenness)
    title : str
        Plot title
    save_path : Path, optional
        If provided, save figure to this path
    top_percent : int
        Keep only top X% of edges for cleaner visualization
    
    Returns:
    --------
    fig : matplotlib.figure.Figure
    """
    conn_vis = conn_matrix.copy()
    
    # Threshold for visualization (keep top edges only)
    if np.any(conn_vis > 0):
        threshold = np.percentile(conn_vis[conn_vis > 0], 100 - top_percent)
        conn_vis[conn_vis < threshold] = 0
    
    # Create graph
    G = nx.from_numpy_array(conn_vis)
    
    # Remove isolated nodes
    isolated = list(nx.isolates(G))
    G.remove_nodes_from(isolated)
    
    if len(G.nodes()) == 0:
        print("  No nodes to display after thresholding")
        return None
    
    n_nodes = len(G.nodes())
    n_edges = len(G.edges())
    
    fig, ax = plt.subplots(figsize=(12, 12))
    
    # Node sizes based on metric
    if node_metric is not None:
        node_metric = np.nan_to_num(node_metric, nan=0)
        # Normalize to reasonable size range
        metric_vals = [node_metric[n] if n < len(node_metric) else 0 for n in G.nodes()]
        if max(metric_vals) > min(metric_vals):
            sizes = 100 + 400 * (np.array(metric_vals) - min(metric_vals)) / \
                    (max(metric_vals) - min(metric_vals))
        else:
            sizes = 200
    else:
        # Size by degree if no metric provided
        degrees = dict(G.degree())
        max_deg = max(degrees.values()) if degrees else 1
        sizes = [100 + 300 * degrees[n] / max_deg for n in G.nodes()]
    
    # Node colors by community
    if community_labels is not None:
        colors = [community_labels[n] if n < len(community_labels) else 0 for n in G.nodes()]
        n_communities = len(np.unique([c for c in colors]))
        cmap = plt.cm.tab20 if n_communities <= 20 else plt.cm.nipy_spectral
    else:
        colors = 'steelblue'
        cmap = None
        n_communities = 1
    
    # Layout - use spring with fixed seed for reproducibility
    pos = nx.spring_layout(G, k=2/np.sqrt(n_nodes), iterations=100, seed=42)
    
    # Edge weights for width
    edges = G.edges()
    weights = np.array([G[u][v]['weight'] for u, v in edges])
    if len(weights) > 0 and weights.max() > 0:
        edge_widths = 0.5 + 2 * weights / weights.max()
    else:
        edge_widths = 1
    
    # Draw edges
    nx.draw_networkx_edges(G, pos, alpha=0.3, width=edge_widths, 
                          edge_color='gray', ax=ax)
    
    # Draw nodes
    nodes = nx.draw_networkx_nodes(G, pos, node_size=sizes, node_color=colors,
                                   cmap=cmap, alpha=0.85, ax=ax)
    
    # Add colorbar for communities if applicable
    if community_labels is not None and cmap is not None:
        plt.colorbar(nodes, ax=ax, label='Community', shrink=0.6)
    
    # Title with info
    info_str = f"Nodes: {n_nodes}, Edges: {n_edges}"
    if community_labels is not None:
        info_str += f", Communities: {n_communities}"
    ax.set_title(f"{title}\n({info_str})", fontsize=12, fontweight='bold')
    ax.axis('off')
    
    plt.tight_layout()
    
    if save_path:
        plt.savefig(save_path, dpi=300, bbox_inches='tight')
        plt.close()
    else:
        plt.show()
    
    return fig


# Cell 7: Main XCP-D Graph Theory Pipeline
class XCPGraphTheoryPipeline:
    """
    Graph theory analysis pipeline for XCP-D outputs.
    
    CORRECTED VERSION v4:
    - New output structure: output_dir/subject/session/{metrics,figures,connectivity_matrices}
    - Individual subject figures
    - Group analysis moved to separate script
    """
    
    def __init__(self, xcp_dir, output_dir, parcellations=None, 
                 use_concatenated=False, thresholds=[0.15], threshold_type='proportional',
                 edge_type='positive_only', expected_parcels=None):
        """
        Initialize the pipeline.
        """
        self.xcp_dir = Path(xcp_dir)
        self.output_dir = Path(output_dir)
        self.parcellations = parcellations
        self.use_concatenated = use_concatenated
        self.thresholds = thresholds
        self.threshold_type = threshold_type
        self.edge_type = edge_type
        self.expected_parcels = expected_parcels or EXPECTED_PARCELS
        
        # Determine primary threshold ONCE (deterministic)
        self.primary_threshold = (PRIMARY_THRESHOLD if PRIMARY_THRESHOLD in self.thresholds 
                                  else self.thresholds[len(self.thresholds)//2])
        
        # Create base output directory only - subject dirs created on the fly
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Results storage
        self.results = {}
        self.summary_df = None
        
    def _get_subject_output_dirs(self, subject_id, session_id):
        """Create and return subject-level output directories."""
        base = self.output_dir / subject_id / session_id
        dirs = {
            'metrics': base / 'metrics',
            'figures': base / 'figures',
            'connectivity_matrices': base / 'connectivity_matrices'
        }
        for d in dirs.values():
            d.mkdir(parents=True, exist_ok=True)
        return dirs
    
    def discover_data(self):
        """Discover all available subjects and parcellations."""
        print("Discovering XCP-D data...")
        print(f"Using per-run files: {not self.use_concatenated}")
        
        subjects = find_xcp_subjects(self.xcp_dir)
        print(f"Found {len(subjects)} subjects")
        
        available_parcellations = set()
        for subj in subjects[:5]:
            for ses_dir in (self.xcp_dir / subj).glob('ses-*'):
                func_dir = ses_dir / 'func'
                if func_dir.exists():
                    parcs = get_available_parcellations(func_dir, self.use_concatenated)
                    available_parcellations.update(parcs)
        
        print(f"Available parcellations: {sorted(available_parcellations)}")
        
        if self.parcellations is None:
            self.parcellations = sorted(available_parcellations)
        else:
            missing = set(self.parcellations) - available_parcellations
            if missing:
                print(f"Warning: Parcellations not found: {missing}")
            self.parcellations = [p for p in self.parcellations if p in available_parcellations]
        
        print(f"Will process: {self.parcellations}")
        print(f"Edge handling: {self.edge_type}")
        print(f"Thresholds: {self.thresholds}")
        print(f"Primary threshold: {self.primary_threshold}")
        
        return subjects
    
    def _extract_run_id(self, filename):
        """Extract run ID from filename (e.g., 'run-01' -> '01')."""
        match = re.search(r'run-(\d+)', str(filename))
        return match.group(1) if match else None
    
    def _generate_subject_figures(self, result, output_dirs):
        """Generate figures for a single subject/session/run."""
        subj = result['subject_id']
        ses = result['session_id']
        run = result['run_id']
        parc = result['parcellation']
        
        run_str = f"_run-{run}" if run else ""
        prefix = f"{subj}_{ses}{run_str}_{parc}"
        fig_dir = output_dirs['figures']
        
        # 1. Raw connectivity matrix
        plot_connectivity_matrix(
            result['connectivity_matrix_raw'],
            title=f'Connectivity (raw r)\n{subj} {ses} run-{run} {parc}',
            save_path=fig_dir / f'{prefix}_connectivity_raw.png'
        )
        
        # 2. Processed connectivity matrix (after edge handling)
        plot_connectivity_matrix(
            result['connectivity_matrix'],
            title=f'Connectivity ({self.edge_type})\n{subj} {ses} run-{run} {parc}',
            save_path=fig_dir / f'{prefix}_connectivity.png'
        )
        
        # 3. Community structure matrix
        if result['primary'] and 'metrics_weighted' in result['primary']:
            mw = result['primary']['metrics_weighted']
            ci = mw.get('community_assignments')
            q = mw.get('modularity', 0)
            if ci is not None:
                plot_community_matrix(
                    result['connectivity_matrix'],
                    ci,
                    title=f'Community Structure\n{subj} {ses} run-{run} {parc} (Q={q:.3f})',
                    save_path=fig_dir / f'{prefix}_community.png'
                )
                
                # 4. Network graph visualization (with communities and strength)
                strength = mw.get('strength')
                plot_network_graph(
                    result['primary']['W'],  # Use thresholded weighted matrix
                    community_labels=ci,
                    node_metric=strength,
                    title=f'Network Graph\n{subj} {ses} run-{run} {parc}',
                    save_path=fig_dir / f'{prefix}_network.png',
                    top_percent=30
                )
        
        # 5. Threshold sweep (if multiple thresholds)
        if len(self.thresholds) > 1 and 'sweep' in result:
            plot_threshold_sweep(
                result['sweep'],
                f'{subj} {ses} run-{run} {parc}',
                save_path=fig_dir / f'{prefix}_threshold_sweep.png'
            )
    
    def process_subject_session(self, ptseries_file, subject_id, session_id, 
                                 parcellation, run_id=None, compute_smallworld=True, n_random=N_RANDOM):
        """
        Process a single ptseries file.
        """
        # Load time series with robust orientation check
        timeseries = load_ptseries(ptseries_file, parcellation=parcellation, 
                                   expected_parcels=self.expected_parcels)
        n_timepoints, n_parcels = timeseries.shape
        
        # Validate shape
        if parcellation in self.expected_parcels:
            expected = self.expected_parcels[parcellation]
            if n_parcels != expected:
                print(f"  WARNING: Expected {expected} parcels for {parcellation}, got {n_parcels}")
                print(f"  File: {ptseries_file}")
                return None
        
        if n_timepoints < 50:
            print(f"  Skipping: only {n_timepoints} timepoints")
            return None
        
        # Compute connectivity matrix (Fisher z-transformed)
        conn_matrix_raw = compute_connectivity_matrix(timeseries, method='pearson', fisher_z=True)
        
        # Prepare graph matrix based on edge handling choice
        conn_matrix = prepare_graph_matrix(conn_matrix_raw, edge_type=self.edge_type)
        
        # === THRESHOLD SWEEP ===
        sweep_results = []
        primary_result = None
        
        for thresh in self.thresholds:
            # Threshold to get weighted matrix
            W = threshold_matrix(conn_matrix, thresh, threshold_type=self.threshold_type, binarize=False)
            
            # Ensure clean matrix for binary conversion
            np.fill_diagonal(W, 0)
            W = (W + W.T) / 2
            
            # Binarize for binary metrics and small-worldness
            A = (W > 0).astype(float)
            np.fill_diagonal(A, 0)
            
            # Compute CORRECT density
            density = compute_density(A)
            
            # Compute weighted metrics
            metrics_w = compute_weighted_metrics(W)
            
            # Compute binary metrics
            metrics_b = compute_binary_metrics(A)
            
            # Compute small-worldness ONLY at primary threshold (expensive)
            sw_metrics = None
            if compute_smallworld and thresh == self.primary_threshold:
                sw_metrics = compute_small_worldness_binary(A, n_random=n_random)
            
            # Store sweep result
            sweep_row = {
                'threshold': thresh,
                'density': density,
                # Weighted metrics
                'global_efficiency_w': metrics_w['global_efficiency'],
                'clustering_mean_w': metrics_w['clustering_mean'],
                'modularity_w': metrics_w['modularity'],
                'n_modules_w': metrics_w['n_modules'],
                'characteristic_path_length_w': metrics_w['characteristic_path_length'],
                'local_efficiency_mean_w': metrics_w['local_efficiency_mean'],
                'transitivity_w': metrics_w['transitivity'],
                'assortativity_w': metrics_w['assortativity'],
                'strength_mean': metrics_w['strength_mean'],
                # Binary metrics
                'global_efficiency_b': metrics_b['global_efficiency'],
                'clustering_mean_b': metrics_b['clustering_mean'],
                'modularity_b': metrics_b['modularity'],
                'characteristic_path_length_b': metrics_b['characteristic_path_length'],
                'degree_mean': metrics_b['degree_mean'],
            }
            
            if sw_metrics:
                sweep_row['small_worldness_sigma'] = sw_metrics['sigma']
                sweep_row['small_worldness_gamma'] = sw_metrics['gamma']
                sweep_row['small_worldness_lambda'] = sw_metrics['lambda']
            
            sweep_results.append(sweep_row)
            
            # Store primary threshold result (deterministic selection)
            if thresh == self.primary_threshold:
                primary_result = {
                    'threshold': thresh,
                    'density': density,
                    'W': W,
                    'A': A,
                    'metrics_weighted': metrics_w,
                    'metrics_binary': metrics_b,
                    'small_worldness': sw_metrics
                }
        
        # Fallback if primary_threshold wasn't in sweep (shouldn't happen)
        if primary_result is None and sweep_results:
            mid_idx = len(self.thresholds) // 2
            thresh = self.thresholds[mid_idx]
            W = threshold_matrix(conn_matrix, thresh, threshold_type=self.threshold_type, binarize=False)
            A = (W > 0).astype(float)
            np.fill_diagonal(A, 0)
            primary_result = {
                'threshold': thresh,
                'density': compute_density(A),
                'W': W,
                'A': A,
                'metrics_weighted': compute_weighted_metrics(W),
                'metrics_binary': compute_binary_metrics(A),
                'small_worldness': compute_small_worldness_binary(A, n_random=n_random) if compute_smallworld else None
            }
        
        return {
            'subject_id': subject_id,
            'session_id': session_id,
            'run_id': run_id,
            'parcellation': parcellation,
            'n_parcels': n_parcels,
            'n_timepoints': n_timepoints,
            'connectivity_matrix_raw': conn_matrix_raw,
            'connectivity_matrix': conn_matrix,
            'primary': primary_result,
            'sweep': pd.DataFrame(sweep_results)
        }
    
    def run(self, compute_smallworld=True, n_random=N_RANDOM, generate_figures=True,
            save_matrices=True):
        """Run the complete pipeline, processing all runs."""
        subjects = self.discover_data()
        
        all_metrics = []
        all_sweep = []
        
        for parcellation in self.parcellations:
            print(f"\n{'='*60}")
            print(f"Processing: {parcellation}")
            print(f"{'='*60}")
            
            parc_results = []
            
            for subject in tqdm(subjects, desc="Subjects"):
                subject_dir = self.xcp_dir / subject
                
                for session in find_sessions(subject_dir):
                    func_dir = subject_dir / session / 'func'
                    
                    if not func_dir.exists():
                        continue
                    
                    ptseries_files = find_ptseries_files(
                        func_dir, parcellation=parcellation, 
                        concatenated=self.use_concatenated
                    )
                    
                    if not ptseries_files:
                        continue
                    
                    # Get output directories for this subject/session
                    output_dirs = self._get_subject_output_dirs(subject, session)
                    
                    # === PROCESS ALL RUNS ===
                    for ptseries_file in ptseries_files:
                        run_id = self._extract_run_id(ptseries_file)
                        run_str = f"run-{run_id}" if run_id else "concatenated"
                        print(f"  Processing: {subject}/{session}/{run_str}")
                        
                        try:
                            result = self.process_subject_session(
                                ptseries_file, subject, session, parcellation,
                                run_id=run_id,
                                compute_smallworld=compute_smallworld, n_random=n_random
                            )
                            
                            if result is not None:
                                parc_results.append(result)
                                
                                # === SAVE CONNECTIVITY MATRICES ===
                                if save_matrices:
                                    run_suffix = f"_run-{run_id}" if run_id else ""
                                    # Save processed matrix
                                    matrix_file = output_dirs['connectivity_matrices'] / \
                                                 f'{subject}_{session}{run_suffix}_{parcellation}_conn.npy'
                                    np.save(matrix_file, result['connectivity_matrix'])
                                    # Save raw matrix too
                                    raw_file = output_dirs['connectivity_matrices'] / \
                                              f'{subject}_{session}{run_suffix}_{parcellation}_conn_raw.npy'
                                    np.save(raw_file, result['connectivity_matrix_raw'])
                                
                                # === GENERATE SUBJECT FIGURES ===
                                if generate_figures:
                                    self._generate_subject_figures(result, output_dirs)
                            
                                # === BUILD METRICS ROW ===
                                p = result['primary']
                                mw = p['metrics_weighted']
                                mb = p['metrics_binary']
                                sw = p['small_worldness']
                                
                                row = {
                                    'subject_id': subject,
                                    'session_id': session,
                                    'run_id': run_id,
                                    'parcellation': parcellation,
                                    'n_parcels': result['n_parcels'],
                                    'n_timepoints': result['n_timepoints'],
                                    'threshold': p['threshold'],
                                    'density': p['density'],
                                    'edge_type': self.edge_type,
                                    # Weighted metrics
                                    'global_efficiency_w': mw['global_efficiency'],
                                    'clustering_mean_w': mw['clustering_mean'],
                                    'modularity_w': mw['modularity'],
                                    'n_modules_w': mw['n_modules'],
                                    'characteristic_path_length_w': mw['characteristic_path_length'],
                                    'local_efficiency_mean_w': mw['local_efficiency_mean'],
                                    'transitivity_w': mw['transitivity'],
                                    'assortativity_w': mw['assortativity'],
                                    'strength_mean': mw['strength_mean'],
                                    'betweenness_mean_w': mw['betweenness_mean'],
                                    'participation_mean_w': mw['participation_mean'],
                                    # Binary metrics
                                    'global_efficiency_b': mb['global_efficiency'],
                                    'clustering_mean_b': mb['clustering_mean'],
                                    'modularity_b': mb['modularity'],
                                    'degree_mean': mb['degree_mean'],
                                    'betweenness_mean_b': mb['betweenness_mean'],
                                    'local_efficiency_mean_b': mb['local_efficiency_mean'],
                                    'assortativity_b': mb['assortativity'],
                                    'transitivity_b': mb['transitivity'],
                                }
                                
                                if sw:
                                    row['small_worldness_sigma'] = sw['sigma']
                                    row['small_worldness_gamma'] = sw['gamma']
                                    row['small_worldness_lambda'] = sw['lambda']
                                    row['C_real'] = sw['C_real']
                                    row['C_rand'] = sw['C_rand']
                                    row['L_real'] = sw['L_real']
                                    row['L_rand'] = sw['L_rand']
                                
                                all_metrics.append(row)
                                
                                # === SAVE METRICS CSV (per-subject) ===
                                row_df = pd.DataFrame([row])
                                run_suffix = f"_run-{run_id}" if run_id else ""
                                metrics_file = output_dirs['metrics'] / \
                                              f'{subject}_{session}{run_suffix}_{parcellation}_metrics.csv'
                                row_df.to_csv(metrics_file, index=False)
                                
                                # === SAVE SWEEP DATA ===
                                sweep_df = result['sweep'].copy()
                                sweep_df['subject_id'] = subject
                                sweep_df['session_id'] = session
                                sweep_df['run_id'] = run_id
                                sweep_df['parcellation'] = parcellation
                                all_sweep.append(sweep_df)
                                
                                sweep_file = output_dirs['metrics'] / \
                                            f'{subject}_{session}{run_suffix}_{parcellation}_threshold_sweep.csv'
                                sweep_df.to_csv(sweep_file, index=False)
                                    
                        except Exception as e:
                            print(f"    Error: {e}")
                            import traceback
                            traceback.print_exc()
                            continue
            
            self.results[parcellation] = parc_results
        
        # Store combined results (for group analysis later)
        self.summary_df = pd.DataFrame(all_metrics)
        self.sweep_df = pd.concat(all_sweep, ignore_index=True) if all_sweep else pd.DataFrame()
        
        # Print summary
        if len(self.summary_df) > 0:
            n_subjects = self.summary_df['subject_id'].nunique()
            n_sessions = self.summary_df.groupby(['subject_id', 'session_id']).ngroups
            n_runs = len(self.summary_df)
            
            print(f"\n{'='*60}")
            print("PROCESSING COMPLETE")
            print(f"{'='*60}")
            print(f"Unique subjects: {n_subjects}")
            print(f"Unique sessions: {n_sessions}")
            print(f"Total runs processed: {n_runs}")
            print(f"Results saved to: {self.output_dir}")
        
        return self.summary_df
    
    def compute_auc_metrics(self):
        """Compute Area Under Curve (AUC) for metrics across threshold sweep."""
        if self.sweep_df is None or len(self.sweep_df) == 0:
            print("No sweep data. Run pipeline first.")
            return None
        
        auc_data = []
        group_cols = ['subject_id', 'session_id', 'run_id', 'parcellation']
        
        for keys, group in self.sweep_df.groupby(group_cols):
            group = group.sort_values('threshold')
            row = dict(zip(group_cols, keys))
            
            metrics_for_auc = ['global_efficiency_w', 'clustering_mean_w', 'modularity_w',
                              'transitivity_w', 'small_worldness_sigma']
            
            for metric in metrics_for_auc:
                if metric in group.columns:
                    data = group[['threshold', metric]].dropna()
                    if len(data) > 1:
                        auc = np.trapezoid(data[metric], data['threshold'])  
                        row[f'{metric}_auc'] = auc
            
            auc_data.append(row)
            
            # Save per-subject AUC
            subj, ses, run, parc = keys
            output_dirs = self._get_subject_output_dirs(subj, ses)
            run_suffix = f"_run-{run}" if run else ""
            auc_file = output_dirs['metrics'] / f'{subj}_{ses}{run_suffix}_{parc}_auc.csv'
            pd.DataFrame([row]).to_csv(auc_file, index=False)
        
        return pd.DataFrame(auc_data)
    
    def compute_session_averages(self):
        """
        Compute session-level averages by averaging across runs within each session.
        Useful for analyses that need one value per session rather than per run.
        """
        if self.summary_df is None or len(self.summary_df) == 0:
            print("No data. Run pipeline first.")
            return None
        
        # Numeric columns to average
        numeric_cols = self.summary_df.select_dtypes(include=[np.number]).columns.tolist()
        # Remove columns that shouldn't be averaged
        exclude = ['n_parcels', 'threshold']
        numeric_cols = [c for c in numeric_cols if c not in exclude]
        
        # Group by subject, session, parcellation and compute mean
        group_cols = ['subject_id', 'session_id', 'parcellation']
        session_avg = self.summary_df.groupby(group_cols)[numeric_cols].mean().reset_index()
        
        # Add metadata
        session_avg['edge_type'] = self.edge_type
        session_avg['threshold'] = self.primary_threshold
        session_avg['n_runs_averaged'] = self.summary_df.groupby(group_cols).size().values
        
        # Save per-subject session averages
        for idx, row in session_avg.iterrows():
            subj = row['subject_id']
            ses = row['session_id']
            parc = row['parcellation']
            output_dirs = self._get_subject_output_dirs(subj, ses)
            avg_file = output_dirs['metrics'] / f'{subj}_{ses}_{parc}_session_avg.csv'
            pd.DataFrame([row]).to_csv(avg_file, index=False)
        
        print(f"Session averages computed: {len(session_avg)} sessions")
        return session_avg
    
    def print_summary(self):
        """Print summary statistics."""
        if self.summary_df is None or len(self.summary_df) == 0:
            print("No results. Run pipeline first.")
            return
        
        print("\n" + "="*70)
        print("GRAPH THEORY ANALYSIS SUMMARY")
        print("="*70)
        print(f"Edge handling: {self.edge_type}")
        print(f"Primary threshold: {self.primary_threshold}")
        print(f"Thresholds tested: {self.thresholds}")
        
        n_subjects = self.summary_df['subject_id'].nunique()
        n_runs = len(self.summary_df)
        print(f"\nTotal: {n_subjects} subjects, {n_runs} runs")
        
        for parc in sorted(self.summary_df['parcellation'].unique()):
            df = self.summary_df[self.summary_df['parcellation'] == parc]
            n_subj = df['subject_id'].nunique()
            n_run = len(df)
            print(f"\n{parc} ({n_subj} subjects, {n_run} runs):")
            print("-"*50)
            
            metrics = [
                ('Global Efficiency (weighted)', 'global_efficiency_w'),
                ('Clustering (weighted)', 'clustering_mean_w'),
                ('Modularity (weighted)', 'modularity_w'),
                ('Small-worldness (σ)', 'small_worldness_sigma'),
                ('Characteristic Path Length (weighted)', 'characteristic_path_length_w'),
                ('Transitivity (weighted)', 'transitivity_w'),
                ('Mean Strength', 'strength_mean'),
                ('Mean Degree (binary)', 'degree_mean'),
                ('Small-worldness AUC (σ)', 'small_worldness_sigma_auc'),
                ('Global Efficiency AUC (weighted)', 'global_efficiency_w_auc'),
                ('Clustering AUC (weighted)', 'clustering_mean_w_auc'),
                ('Modularity AUC (weighted)', 'modularity_w_auc'),
                ('Transitivity AUC (weighted)', 'transitivity_w_auc'),
            ]
            
            for name, col in metrics:
                if col in df.columns:
                    data = df[col].replace([np.inf, -np.inf], np.nan).dropna()
                    if len(data) > 0:
                        print(f"  {name:35s}: {data.mean():.4f} ± {data.std():.4f}")
        
        print("\n" + "="*70)

# Add this at the END of the file (after the XCPGraphTheoryPipeline class)

# =============================================================================
# MAIN ENTRY POINT
# =============================================================================
def main():
    parser = argparse.ArgumentParser(
        description='Graph Theory Analysis Pipeline for XCP-D outputs',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python run_gt.py --parcellation Glasser
    python run_gt.py --parcellation Glasser Gordon 4S456Parcels
    python run_gt.py --all
    python run_gt.py --parcellation Glasser --no-smallworld  # Skip small-worldness (faster)
        """
    )
    parser.add_argument('--parcellation', '-p', nargs='+', type=str,
                        help='Parcellation(s) to process')
    parser.add_argument('--all', '-a', action='store_true',
                        help='Process all available parcellations')
    parser.add_argument('--no-smallworld', action='store_true',
                        help='Skip small-worldness computation (much faster)')
    parser.add_argument('--n-random', type=int, default=100,
                        help='Number of random networks for small-worldness (default: 100)')
    parser.add_argument('--no-figures', action='store_true',
                        help='Skip figure generation')
    parser.add_argument('--save-matrices', action='store_true',
                        help='Save connectivity matrices as .npy files')
    parser.add_argument('--output-dir', '-o', type=str, default=OUTPUT_DIR,
                        help=f'Output directory (default: {OUTPUT_DIR})')
    
    args = parser.parse_args()
    
    # Validate arguments
    if not args.parcellation and not args.all:
        parser.error("Must specify --parcellation or --all")
    
    # Determine parcellations to process
    parcellations = None if args.all else args.parcellation
    
    print("\n" + "="*70)
    print("GRAPH THEORY ANALYSIS PIPELINE")
    print("="*70)
    print(f"Output directory: {args.output_dir}")
    print(f"Parcellations: {'ALL' if args.all else args.parcellation}")
    print(f"Compute small-worldness: {not args.no_smallworld}")
    print(f"N random networks: {args.n_random}")
    print("="*70 + "\n")
    
    # Initialize and run pipeline
    pipeline = XCPGraphTheoryPipeline(
        xcp_dir=XCP_OUTPUT_DIR,
        output_dir=args.output_dir,
        parcellations=parcellations,
        use_concatenated=USE_CONCATENATED,
        thresholds=THRESHOLDS,
        threshold_type=THRESHOLD_TYPE,
        edge_type=EDGE_TYPE
    )
    
    summary_df = pipeline.run(
        compute_smallworld=not args.no_smallworld,
        n_random=args.n_random,
        generate_figures=not args.no_figures,
        save_matrices=args.save_matrices
    )
    
    # Compute additional metrics
    pipeline.compute_auc_metrics()
    pipeline.compute_session_averages()
    pipeline.print_summary()
    
    print(f"\nAll results saved to: {args.output_dir}")
    return summary_df


if __name__ == '__main__':
    main()