"""Report configuration models."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import List, Optional

import pandas as pd


class ReportType(Enum):
    """Type of report to generate."""
    QUARTERLY = "quarterly"           # Standard quarterly report (Q1-Q4)
    CUSTOM_PERIOD = "custom_period"   # Custom date range
    MONTHLY = "monthly"               # Single month report
    YEARLY = "yearly"                 # Full year report
    PROJECT_SUMMARY = "project"       # Specific projects only
    EMPLOYEE_SUMMARY = "employee"     # Specific employees only


class TimeGrouping(Enum):
    """How to group data temporally in the report."""
    BY_MONTH = "monthly"              # Separate blocks per month (current behavior)
    BY_PERIOD = "period"              # Single block for entire period
    BY_WEEK = "weekly"                # Weekly blocks
    NONE = "none"                     # No time grouping, total sum only


@dataclass
class ReportConfig:
    """Configuration for flexible report generation."""

    report_type: ReportType
    start_date: date
    end_date: date
    time_grouping: TimeGrouping

    # Optional filters
    projects: Optional[List[str]] = None        # Filter to specific projects
    employees: Optional[List[str]] = None       # Filter to specific employees

    # Report components to include
    include_bonus_calc: bool = True             # Include bonus calculations
    include_budget_overview: bool = True        # Include budget overview sheet
    include_summary_sheet: bool = True          # Include summary cover sheet
    include_quarterly_summary: bool = True      # Include quarterly summary section

    # Additional options
    exclude_special_projects: bool = False      # Exclude 0000/0.1000 projects

    def __post_init__(self):
        """Validate configuration."""
        if self.start_date > self.end_date:
            raise ValueError("start_date must be before end_date")

        # For quarterly reports, validate date range spans exactly one quarter
        if self.report_type == ReportType.QUARTERLY:
            delta = (self.end_date - self.start_date).days
            if delta < 85 or delta > 95:  # ~3 months
                raise ValueError("Quarterly report must span approximately 3 months")


@dataclass
class TimeBlock:
    """Represents a time period for grouping report data."""

    name: str                    # Display name (e.g., "Juli 2025" or "15.08-15.09")
    start: date                  # Block start date
    end: date                    # Block end date
    data: pd.DataFrame           # Filtered data for this time block
    period: Optional[pd.Period] = None  # Optional pandas Period for month-based blocks

    def __str__(self) -> str:
        return self.name

    @property
    def duration_days(self) -> int:
        """Number of days in this time block."""
        return (self.end - self.start).days + 1
