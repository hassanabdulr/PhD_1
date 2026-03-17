#!/usr/bin/env python
"""
Group-level analysis and figures for Graph Theory Pipeline.

Run this AFTER individual subject processing is complete.

Usage:
    python group_analysis.py --input-dir /path/to/output --parcellation Glasser
    python group_analysis.py --input-dir /path/to/output --all
"""

import argparse
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path
import bct
from glob import glob

# Import from main script
from run_gt import (
    plot_connectivity_matrix, plot_community_matrix, plot_metrics_summary,
    threshold_matrix, EXPECTED_PARCELS, PRIMARY_THRESHOLD
)

plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("husl")


def collect_metrics(input_dir, parcellation=None):
    """Collect all metrics CSVs from subject directories."""
    input_path = Path(input_dir)
    
    # Find all metrics files
    pattern = f"**/metrics/*_{parcellation}_metrics.csv" if parcellation else "**/metrics/*_metrics.csv"
    files = list(input_path.glob(pattern))
    
    if not files:
        print(f"No metrics files found for pattern: {pattern}")
        return None
    
    print(f"Found {len(files)} metrics files")
    
    dfs = []
    for f in files:
        try:
            df = pd.read_csv(f)
            dfs.append(df)
        except Exception as e:
            print(f"Error reading {f}: {e}")
    
    if not dfs:
        return None
    
    combined = pd.concat(dfs, ignore_index=True)
    print(f"Combined metrics: {len(combined)} rows")
    return combined


def collect_connectivity_matrices(input_dir, parcellation):
    """Collect all connectivity matrices for a parcellation."""
    input_path = Path(input_dir)
    
    # Find processed (not raw) matrices
    pattern = f"**/connectivity_matrices/*_{parcellation}_conn.npy"
    files = list(input_path.glob(pattern))
    # Exclude raw matrices
    files = [f for f in files if '_raw.npy' not in str(f)]
    
    if not files:
        print(f"No connectivity matrices found for {parcellation}")
        return None, []
    
    print(f"Found {len(files)} connectivity matrices for {parcellation}")
    
    matrices = []
    labels = []
    for f in files:
        try:
            mat = np.load(f)
            matrices.append(mat)
            labels.append(f.stem.replace('_conn', ''))
        except Exception as e:
            print(f"Error loading {f}: {e}")
    
    if not matrices:
        return None, []
    
    # Stack into 3D array
    matrices = np.stack(matrices, axis=0)
    return matrices, labels


def generate_group_figures(input_dir, output_dir, parcellation, threshold=PRIMARY_THRESHOLD):
    """Generate group-level figures for a parcellation."""
    
    # Collect data
    matrices, labels = collect_connectivity_matrices(input_dir, parcellation)
    metrics_df = collect_metrics(input_dir, parcellation)
    
    if matrices is None or metrics_df is None:
        print(f"Insufficient data for {parcellation}")
        return
    
    # Create output directory
    group_dir = Path(output_dir) / 'group' / parcellation
    group_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"\nGenerating group figures for {parcellation}...")
    print(f"  N subjects/runs: {len(matrices)}")
    
    # 1. Group average connectivity matrix
    avg_matrix = np.mean(matrices, axis=0)
    plot_connectivity_matrix(
        avg_matrix,
        title=f'Group Average Connectivity - {parcellation}\n(n={len(matrices)})',
        save_path=group_dir / 'group_avg_connectivity.png'
    )
    
    # 2. Group community structure
    avg_thresh = threshold_matrix(avg_matrix, threshold, threshold_type='proportional')
    
    # Best-of-10 Louvain
    best_q = -np.inf
    best_ci = None
    for _ in range(10):
        ci_temp, q_temp = bct.community_louvain(avg_thresh)
        if q_temp > best_q:
            best_q = q_temp
            best_ci = ci_temp
    
    plot_community_matrix(
        avg_matrix, best_ci,
        title=f'Group Community Structure - {parcellation}\n(Q={best_q:.3f}, {len(np.unique(best_ci))} modules)',
        save_path=group_dir / 'group_community_structure.png'
    )
    
    # 3. Metrics distribution
    plot_metrics_summary(
        metrics_df, parcellation,
        save_path=group_dir / 'metrics_distribution.png'
    )
    
    # 4. Variance map (optional - shows inter-subject variability)
    std_matrix = np.std(matrices, axis=0)
    fig, ax = plt.subplots(figsize=(10, 8))
    im = ax.imshow(std_matrix, cmap='hot', aspect='equal')
    plt.colorbar(im, ax=ax, shrink=0.8, label='Std Dev')
    ax.set_title(f'Inter-subject Variability - {parcellation}\n(n={len(matrices)})')
    ax.set_xlabel('Parcels')
    ax.set_ylabel('Parcels')
    plt.tight_layout()
    plt.savefig(group_dir / 'connectivity_variability.png', dpi=300, bbox_inches='tight')
    plt.close()
    
    # 5. Save combined metrics
    metrics_df.to_csv(group_dir / f'graph_metrics_{parcellation}.csv', index=False)
    
    # 6. Save group average matrix
    np.save(group_dir / f'group_avg_connectivity_{parcellation}.npy', avg_matrix)
    
    print(f"  Figures saved to: {group_dir}")


def main():
    parser = argparse.ArgumentParser(description='Group-level analysis for Graph Theory Pipeline')
    parser.add_argument('--input-dir', '-i', type=str, required=True,
                        help='Input directory (output from run_gt.py)')
    parser.add_argument('--output-dir', '-o', type=str, default=None,
                        help='Output directory (default: same as input)')
    parser.add_argument('--parcellation', '-p', nargs='+', type=str,
                        help='Parcellation(s) to process')
    parser.add_argument('--all', '-a', action='store_true',
                        help='Process all available parcellations')
    parser.add_argument('--threshold', '-t', type=float, default=PRIMARY_THRESHOLD,
                        help=f'Threshold for group analysis (default: {PRIMARY_THRESHOLD})')
    
    args = parser.parse_args()
    
    output_dir = args.output_dir or args.input_dir
    
    # Discover available parcellations if --all
    if args.all:
        input_path = Path(args.input_dir)
        files = list(input_path.glob('**/metrics/*_metrics.csv'))
        parcellations = set()
        for f in files:
            # Extract parcellation from filename like sub-001_ses-01_run-01_Glasser_metrics.csv
            parts = f.stem.replace('_metrics', '').split('_')
            if len(parts) >= 4:
                parcellations.add(parts[-1])
        parcellations = sorted(parcellations)
        print(f"Found parcellations: {parcellations}")
    else:
        parcellations = args.parcellation
    
    if not parcellations:
        print("No parcellations specified or found")
        return
    
    for parc in parcellations:
        generate_group_figures(args.input_dir, output_dir, parc, args.threshold)
    
    print(f"\nGroup analysis complete. Results in: {output_dir}/group/")


if __name__ == '__main__':
    main()