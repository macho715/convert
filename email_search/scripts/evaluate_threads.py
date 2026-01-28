#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Evaluate thread clustering accuracy against a pairwise ground truth JSON.
Ground Truth Format (List of dicts):
[
  {"id_a": 1, "id_b": 2, "match": true, ...},
  ...
]
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Set, Tuple

def load_json(path: Path) -> List[dict]:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))

def build_pred_lookup(threads: List[dict]) -> Dict[int, str]:
    """
    Build a lookup map: row_index -> thread_id
    """
    lookup = {}
    for t in threads:
        tid = t["thread_id"]
        for m in t["members"]:
            lookup[int(m)] = tid
    return lookup

def evaluate(pred_path: Path, gt_path: Path):
    print(f"Loading Prediction: {pred_path}")
    preds = load_json(pred_path)
    pred_lookup = build_pred_lookup(preds)
    
    print(f"Loading Ground Truth: {gt_path}")
    gts = load_json(gt_path)
    
    tp = 0
    tn = 0
    fp = 0
    fn = 0
    
    for item in gts:
        id_a = int(item["id_a"])
        id_b = int(item["id_b"])
        is_match_gt = item["match"]
        
        tid_a = pred_lookup.get(id_a)
        tid_b = pred_lookup.get(id_b)
        
        # Prediction: Match if both have same thread_id (and not None)
        is_match_pred = (tid_a is not None) and (tid_b is not None) and (tid_a == tid_b)
        
        if is_match_gt and is_match_pred:
            tp += 1
        elif not is_match_gt and not is_match_pred:
            tn += 1
        elif not is_match_gt and is_match_pred:
            fp += 1
        elif is_match_gt and not is_match_pred:
            fn += 1
            
    total = tp + tn + fp + fn
    if total == 0:
        print("No ground truth pairs found.")
        return

    accuracy = (tp + tn) / total
    precision = tp / (tp + fp) if (tp + fp) > 0 else 0.0
    recall = tp / (tp + fn) if (tp + fn) > 0 else 0.0
    f1 = 2 * (precision * recall) / (precision + recall) if (precision + recall) > 0 else 0.0
    
    print("-" * 40)
    print(f"Total Pairs Evaluated: {total}")
    print("-" * 40)
    print(f"True Positives (TP): {tp}")
    print(f"True Negatives (TN): {tn}")
    print(f"False Positives (FP): {fp}")
    print(f"False Negatives (FN): {fn}")
    print("-" * 40)
    print(f"Accuracy:  {accuracy:.4f}")
    print(f"Precision: {precision:.4f}")
    print(f"Recall:    {recall:.4f}")
    print(f"F1 Score:  {f1:.4f}")
    print("-" * 40)

def main():
    parser = argparse.ArgumentParser(description="Evaluate Thread Clustering")
    parser.add_argument("pred", help="Prediction JSON (threads.json)")
    parser.add_argument("gt", help="Ground Truth JSON (evaluation_set.json)")
    args = parser.parse_args()
    
    evaluate(Path(args.pred), Path(args.gt))

if __name__ == "__main__":
    main()
