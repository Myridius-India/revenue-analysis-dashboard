from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from src.revenue_loader import build_project_snapshot, load_master_projects, load_revenue_files, merge_with_master, summarize


@dataclass(frozen=True)
class RevenueDataset:
    merged: pd.DataFrame
    snapshot: pd.DataFrame
    summary: pd.DataFrame
    revenue_count: int
    master_count: int


def load_local_dataset(data_folder: Path, template_file: Path) -> RevenueDataset:
    revenue_data = load_revenue_files(Path(data_folder))
    master_data = load_master_projects(Path(template_file))
    merged = merge_with_master(revenue_data, master_data)
    snapshot = build_project_snapshot(merged)
    summary = summarize(merged, "Project")
    return RevenueDataset(
        merged=merged,
        snapshot=snapshot,
        summary=summary,
        revenue_count=len(revenue_data),
        master_count=len(master_data),
    )
