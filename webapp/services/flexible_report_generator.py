"""Flexible report generator supporting various time groupings and filters."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, List, Optional

import pandas as pd

from ..models import ReportConfig, ReportType, TimeGrouping, TimeBlock
from ..report_generator import (
    build_quarterly_report,
    load_csv_budget_data,
    load_csv_projects,
    load_xml_times,
    determine_quarter,
    MONTH_NAMES,
)
from .flexible_report_builder import build_flexible_report


ProgressCallback = Callable[[int, str], None]


class FlexibleReportGenerator:
    """
    Generates reports with flexible time grouping and filtering.

    Supports:
    - Custom date ranges (e.g., 15.08-15.09)
    - Different time groupings (by month, by period, by week, none)
    - Project and employee filtering
    - Configurable report components
    """

    def __init__(
        self,
        config: ReportConfig,
        csv_path: Path,
        xml_path: Path,
        progress_cb: ProgressCallback = lambda p, m: None,
    ):
        """
        Initialize flexible report generator.

        Args:
            config: Report configuration
            csv_path: Path to CSV budget file
            xml_path: Path to XML timesheet data
            progress_cb: Optional callback for progress updates
        """
        self.config = config
        self.csv_path = csv_path
        self.xml_path = xml_path
        self.progress_cb = progress_cb

    def generate(self, output_path: Path) -> Path:
        """
        Generate report based on configuration.

        Args:
            output_path: Path where to save the generated Excel file

        Returns:
            Path to the generated Excel file
        """
        self.progress_cb(5, "Lade Daten...")

        # Load data
        df_csv = load_csv_projects(self.csv_path)
        df_budget, milestone_parent_map = load_csv_budget_data(self.csv_path)
        df_xml = load_xml_times(self.xml_path)

        self.progress_cb(15, "Filtere und gruppiere Daten...")

        # Filter data by date range
        df_xml_filtered = self._filter_by_date_range(df_xml)

        # Apply project filter if specified
        if self.config.projects:
            df_xml_filtered = self._filter_by_projects(df_xml_filtered)

        # Apply employee filter if specified
        if self.config.employees:
            df_xml_filtered = self._filter_by_employees(df_xml_filtered)

        # Exclude special projects if requested
        if self.config.exclude_special_projects:
            df_xml_filtered = self._exclude_special_projects(df_xml_filtered)

        # Create time blocks based on grouping strategy
        time_blocks = self._create_time_blocks(df_xml_filtered)

        self.progress_cb(20, f"Erstelle Report mit {len(time_blocks)} Zeitblöcken...")

        # For quarterly reports with monthly grouping, use existing logic
        if (self.config.report_type == ReportType.QUARTERLY and
            self.config.time_grouping == TimeGrouping.BY_MONTH):
            return self._generate_quarterly_report(
                df_csv, df_budget, milestone_parent_map, df_xml_filtered, output_path
            )

        # For custom configurations, use new flexible logic
        return self._generate_flexible_report(
            df_csv, df_budget, milestone_parent_map, df_xml_filtered,
            time_blocks, output_path
        )

    def _filter_by_date_range(self, df_xml: pd.DataFrame) -> pd.DataFrame:
        """Filter XML data to specified date range."""
        # Use date_parsed column which is already datetime from load_xml_times
        if 'date_parsed' not in df_xml.columns:
            raise ValueError("XML data must have a 'date_parsed' column")

        # Filter by date range
        mask = (
            (df_xml['date_parsed'].dt.date >= self.config.start_date) &
            (df_xml['date_parsed'].dt.date <= self.config.end_date)
        )

        return df_xml[mask].copy()

    def _filter_by_projects(self, df_xml: pd.DataFrame) -> pd.DataFrame:
        """Filter to specified projects only."""
        if not self.config.projects:
            return df_xml

        # Match by proj_norm (normalized project code)
        mask = df_xml['proj_norm'].isin(self.config.projects)

        # Also try matching by project code prefix
        for proj in self.config.projects:
            mask |= df_xml['proj_norm'].str.startswith(proj)

        return df_xml[mask].copy()

    def _filter_by_employees(self, df_xml: pd.DataFrame) -> pd.DataFrame:
        """Filter to specified employees only."""
        if not self.config.employees:
            return df_xml

        mask = df_xml['staff_name'].isin(self.config.employees)
        return df_xml[mask].copy()

    def _exclude_special_projects(self, df_xml: pd.DataFrame) -> pd.DataFrame:
        """Exclude special projects (0000, 0.1000)."""
        mask = ~df_xml['proj_norm'].str.match(r'^0\.?0+')
        return df_xml[mask].copy()

    def _create_time_blocks(self, df_xml: pd.DataFrame) -> List[TimeBlock]:
        """
        Create time blocks based on configured grouping strategy.

        Args:
            df_xml: Filtered XML timesheet data

        Returns:
            List of TimeBlock objects
        """
        if self.config.time_grouping == TimeGrouping.BY_MONTH:
            return self._group_by_month(df_xml)

        elif self.config.time_grouping == TimeGrouping.BY_PERIOD:
            return self._group_by_period(df_xml)

        elif self.config.time_grouping == TimeGrouping.BY_WEEK:
            return self._group_by_week(df_xml)

        else:  # TimeGrouping.NONE
            return self._group_total(df_xml)

    def _group_by_month(self, df_xml: pd.DataFrame) -> List[TimeBlock]:
        """Group data by calendar months."""
        blocks = []

        # Use date_parsed which is already datetime from load_xml_times
        # Get unique months in the data
        df_xml['month_period'] = df_xml['date_parsed'].dt.to_period('M')
        months = sorted(df_xml['month_period'].unique())

        for month in months:
            month_data = df_xml[df_xml['month_period'] == month].copy()

            # German month name
            month_name = MONTH_NAMES.get(month.month, month.strftime('%B'))
            name = f"{month_name} {month.year}"

            # Get first and last day of month
            start = month.to_timestamp().date()
            end = (month + 1).to_timestamp().date() - timedelta(days=1)

            blocks.append(TimeBlock(
                name=name,
                start=start,
                end=end,
                data=month_data,
                period=month,
            ))

        return blocks

    def _group_by_period(self, df_xml: pd.DataFrame) -> List[TimeBlock]:
        """Create a single block for the entire period."""
        name = f"{self.config.start_date.strftime('%d.%m.%Y')} - {self.config.end_date.strftime('%d.%m.%Y')}"

        return [TimeBlock(
            name=name,
            start=self.config.start_date,
            end=self.config.end_date,
            data=df_xml.copy(),
            period=None,
        )]

    def _group_by_week(self, df_xml: pd.DataFrame) -> List[TimeBlock]:
        """Group data by calendar weeks."""
        blocks = []

        # Use date_parsed which is already datetime from load_xml_times
        # Get unique weeks
        df_xml['week'] = df_xml['date_parsed'].dt.isocalendar().week
        df_xml['year'] = df_xml['date_parsed'].dt.year

        weeks = df_xml.groupby(['year', 'week']).groups

        for (year, week), indices in sorted(weeks.items()):
            week_data = df_xml.loc[indices].copy()

            # Get start and end of week
            start = week_data['date_parsed'].min().date()
            end = week_data['date_parsed'].max().date()

            name = f"KW {week} ({start.strftime('%d.%m')} - {end.strftime('%d.%m.%Y')})"

            blocks.append(TimeBlock(
                name=name,
                start=start,
                end=end,
                data=week_data,
                period=None,
            ))

        return blocks

    def _group_total(self, df_xml: pd.DataFrame) -> List[TimeBlock]:
        """Create a single total block with no time subdivision."""
        return [TimeBlock(
            name="Gesamt",
            start=self.config.start_date,
            end=self.config.end_date,
            data=df_xml.copy(),
            period=None,
        )]

    def _generate_quarterly_report(
        self,
        df_csv: pd.DataFrame,
        df_budget: pd.DataFrame,
        milestone_parent_map,
        df_xml: pd.DataFrame,
        output_path: Path,
    ) -> Path:
        """Generate standard quarterly report using existing logic."""

        # Determine quarter from start date
        quarter = pd.Period(self.config.start_date, freq="Q")

        # Get months in quarter
        months = pd.period_range(
            start=quarter.start_time,
            end=quarter.end_time,
            freq='M'
        )
        
        report_title = f"Quartalsübersicht {quarter}"

        # Use existing build function
        return build_quarterly_report(
            df_csv=df_csv,
            df_budget=df_budget,
            milestone_parent_map=milestone_parent_map,
            df_xml=df_xml, # This is already filtered to the quarter
            target_quarter=quarter,
            months=months,
            out_path=output_path,
            progress_cb=self.progress_cb,
            add_vba=True,
            report_title=report_title,
            use_quarter_filter=False, # Data is already filtered
        )

    def _generate_flexible_report(
        self,
        df_csv: pd.DataFrame,
        df_budget: pd.DataFrame,
        milestone_parent_map,
        df_xml: pd.DataFrame,
        time_blocks: List[TimeBlock],
        output_path: Path,
    ) -> Path:
        """
        Generate report with flexible time blocks using the new builder.
        """
        # The new builder handles everything from here
        return build_flexible_report(
            config=self.config,
            df_csv=df_csv,
            df_budget=df_budget,
            milestone_parent_map=milestone_parent_map,
            time_blocks=time_blocks,
            out_path=output_path,
            progress_cb=self.progress_cb,
        )
