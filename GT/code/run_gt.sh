#!/bin/bash
#SBATCH --job-name=GT_analysis
#SBATCH --partition=high-moby
#SBATCH --output=/projects/aabdulrasul/TAY/GT/logs/gt_%A_%a.out
#SBATCH --error=/projects/aabdulrasul/TAY/GT/logs/gt_%A_%a.err
#SBATCH --array=0-15
#SBATCH --cpus-per-task=4
#SBATCH --mem=16G
#SBATCH --time=12:00:00

mkdir -p /projects/aabdulrasul/TAY/GT/logs

source /projects/aabdulrasul/TAY/GT/code/gt_env/bin/activate


PARCELLATIONS=(
    '4S156Parcels' '4S256Parcels' '4S356Parcels'
    '4S456Parcels' '4S556Parcels' '4S656Parcels'
    '4S756Parcels' '4S856Parcels' '4S956Parcels'
    '4S1056Parcels' 'Glasser' 'Gordon'
    'HCP' 'MIDB' 'MyersLabonte' 'Tian'
)
PARC=${PARCELLATIONS[$SLURM_ARRAY_TASK_ID]}

echo "Processing: $PARC"
python /projects/aabdulrasul/TAY/GT/code/run_gt.py --parcellation "$PARC" --save-matrices