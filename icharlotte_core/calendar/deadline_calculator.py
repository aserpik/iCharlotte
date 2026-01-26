"""
California Civil Litigation Deadline Calculator.

Computes legal deadlines based on rules in legal_rules_manifest.json,
handling court days, calendar days, service extensions, and holidays.
"""

import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple

from icharlotte_core.ui.logs_tab import LogManager


# Path to the rules manifest
RULES_FILE = Path(__file__).parent / 'legal_rules_manifest.json'


class DeadlineCalculator:
    """
    California civil litigation deadline calculator.

    Uses legal_rules_manifest.json to compute deadlines with proper
    handling of court days, calendar days, service extensions, and holidays.

    Usage:
        calc = DeadlineCalculator()

        # Calculate MSJ opposition deadline
        deadline = calc.calculate_deadline(
            rule_slug='msj-opposition',
            trigger_date=datetime(2026, 3, 20),  # hearing date
            service_method='electronic'
        )

        # Get all deadlines for a motion
        deadlines = calc.get_motion_deadlines(
            motion_type='msj',
            hearing_date=datetime(2026, 3, 20),
            service_method='electronic'
        )
    """

    def __init__(self):
        """Initialize the calculator and load rules."""
        self.log = LogManager()
        self.rules_data = self._load_rules()
        self.rules_by_slug = {
            rule['rule_slug']: rule
            for rule in self.rules_data.get('rules', [])
        }
        self.holidays = self._build_holiday_set()

    def _load_rules(self) -> Dict:
        """Load the legal rules manifest."""
        try:
            with open(RULES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            self.log.add_log("Calendar", f"Failed to load rules manifest: {e}")
            return {}

    def _build_holiday_set(self) -> set:
        """Build a set of holiday dates for quick lookup."""
        holidays = set()
        holiday_dates = self.rules_data.get('holidays', {}).get('holiday_dates', {})

        for year, dates in holiday_dates.items():
            for entry in dates:
                try:
                    date_str = entry.get('date', '')
                    date = datetime.strptime(date_str, '%Y-%m-%d').date()
                    holidays.add(date)
                except:
                    pass

        return holidays

    def is_court_day(self, date: datetime) -> bool:
        """
        Check if a date is a court day (not weekend or holiday).

        Args:
            date: The date to check

        Returns:
            True if it's a court day, False otherwise
        """
        d = date.date() if isinstance(date, datetime) else date

        # Check weekend
        if d.weekday() >= 5:  # Saturday = 5, Sunday = 6
            return False

        # Check holidays
        if d in self.holidays:
            return False

        return True

    def count_court_days(self, start_date: datetime, num_days: int, direction: str = 'forward') -> datetime:
        """
        Count court days from a start date.

        Args:
            start_date: The starting date (excluded from count)
            num_days: Number of court days to count
            direction: 'forward' or 'backward'

        Returns:
            The resulting date after counting court days
        """
        if num_days == 0:
            return start_date

        current = start_date
        days_counted = 0
        step = 1 if direction == 'forward' else -1

        while days_counted < num_days:
            current = current + timedelta(days=step)
            if self.is_court_day(current):
                days_counted += 1

        return current

    def count_calendar_days(self, start_date: datetime, num_days: int, direction: str = 'forward') -> datetime:
        """
        Count calendar days from a start date.

        Args:
            start_date: The starting date (excluded from count)
            num_days: Number of calendar days to count
            direction: 'forward' or 'backward'

        Returns:
            The resulting date after counting calendar days
        """
        step = 1 if direction == 'forward' else -1
        return start_date + timedelta(days=num_days * step)

    def apply_ccp_12a_adjustment(self, date: datetime, direction: str) -> datetime:
        """
        Apply CCP § 12a holiday adjustment.

        If the computed date falls on a holiday/weekend:
        - For backward-counting: move to preceding court day
        - For forward-counting: move to following court day

        Args:
            date: The computed deadline date
            direction: 'backward' or 'forward' (the original counting direction)

        Returns:
            The adjusted date
        """
        if self.is_court_day(date):
            return date

        # Determine adjustment direction based on original counting direction
        # Backward-counting deadlines: must file EARLIER (move to preceding court day)
        # Forward-counting deadlines: get MORE time (move to following court day)
        if direction == 'backward':
            # Move to preceding court day
            while not self.is_court_day(date):
                date = date - timedelta(days=1)
        else:
            # Move to following court day
            while not self.is_court_day(date):
                date = date + timedelta(days=1)

        return date

    def calculate_deadline(
        self,
        rule_slug: str,
        trigger_date: datetime,
        service_method: str = 'electronic'
    ) -> Optional[datetime]:
        """
        Calculate a deadline based on a rule and trigger date.

        Args:
            rule_slug: The rule identifier (e.g., 'msj-opposition')
            trigger_date: The trigger date (e.g., hearing date)
            service_method: Service method for extension calculation

        Returns:
            The calculated deadline date, or None if rule not found
        """
        rule = self.rules_by_slug.get(rule_slug)
        if not rule:
            self.log.add_log("Calendar", f"Rule not found: {rule_slug}")
            return None

        logic = rule.get('logic', {})
        offset = logic.get('offset', 0)
        unit = logic.get('unit', 'calendar_days')
        direction = logic.get('direction', 'backward')

        # Step 1: Calculate base deadline
        if unit == 'court_days':
            base_deadline = self.count_court_days(trigger_date, offset, direction)
        else:  # calendar_days
            base_deadline = self.count_calendar_days(trigger_date, offset, direction)

        # Step 2: Apply service extension if applicable
        service_extensions = rule.get('service_extensions', {})
        extension = service_extensions.get(service_method, {})

        if extension:
            ext_days = extension.get('days', 0)
            ext_unit = extension.get('unit', 'calendar_days')

            # Service extension always adds time (moves deadline forward/later)
            if ext_unit == 'court_days':
                base_deadline = self.count_court_days(base_deadline, ext_days, 'forward')
            else:
                base_deadline = self.count_calendar_days(base_deadline, ext_days, 'forward')

        # Step 3: Apply CCP § 12a adjustment
        final_deadline = self.apply_ccp_12a_adjustment(base_deadline, direction)

        return final_deadline

    def get_rule_info(self, rule_slug: str) -> Optional[Dict]:
        """Get full rule information by slug."""
        return self.rules_by_slug.get(rule_slug)

    def get_motion_deadlines(
        self,
        motion_type: str,
        hearing_date: datetime,
        service_method: str = 'electronic'
    ) -> List[Dict[str, Any]]:
        """
        Get all deadlines for a motion type (notice, opposition, reply).

        Args:
            motion_type: Motion type ('msj', 'msa', 'standard', 'demurrer', etc.)
            hearing_date: The hearing date
            service_method: Service method for extensions

        Returns:
            List of deadline dicts with 'title', 'date', 'description', 'rule_slug'
        """
        deadlines = []

        # Map motion types to rule prefixes
        motion_prefixes = {
            'msj': 'msj',
            'msa': 'msa',
            'standard': 'standard-motion',
            'demurrer': 'demurrer',
            'discovery': 'discovery-motion',
            'motion-to-strike': 'motion-to-strike',
        }

        prefix = motion_prefixes.get(motion_type.lower(), 'standard-motion')

        # Find all related rules
        deadline_types = [
            ('notice', 'Notice of Motion Due'),
            ('opposition', 'Opposition Due'),
            ('reply', 'Reply Due'),
        ]

        for suffix, title in deadline_types:
            rule_slug = f"{prefix}-{suffix}"
            rule = self.rules_by_slug.get(rule_slug)

            if rule:
                deadline_date = self.calculate_deadline(
                    rule_slug=rule_slug,
                    trigger_date=hearing_date,
                    service_method=service_method
                )

                if deadline_date:
                    deadlines.append({
                        'title': title,
                        'date': deadline_date,
                        'description': f"{rule.get('name', '')}\nStatute: {rule.get('statute', '')}\nService: {service_method}",
                        'rule_slug': rule_slug,
                    })

        return deadlines

    def get_discovery_response_deadline(
        self,
        request_date: datetime,
        service_method: str = 'electronic'
    ) -> Dict[str, Any]:
        """
        Calculate discovery response deadline (30 calendar days + service extension).

        Args:
            request_date: Date discovery request was served
            service_method: Service method

        Returns:
            Deadline dict with 'title', 'date', 'description'
        """
        # Base: 30 calendar days from request
        base_deadline = self.count_calendar_days(request_date, 30, 'forward')

        # Apply service extension
        if service_method == 'electronic':
            base_deadline = self.count_court_days(base_deadline, 2, 'forward')
        elif service_method == 'mail_ca':
            base_deadline = self.count_calendar_days(base_deadline, 5, 'forward')
        elif service_method == 'mail_other_state':
            base_deadline = self.count_calendar_days(base_deadline, 10, 'forward')
        elif service_method == 'overnight':
            base_deadline = self.count_calendar_days(base_deadline, 2, 'forward')

        # Apply CCP 12a
        final_deadline = self.apply_ccp_12a_adjustment(base_deadline, 'forward')

        return {
            'title': 'Discovery Response Due',
            'date': final_deadline,
            'description': f"Response due 30 calendar days + {service_method} service extension\nCCP § 2030.260, 2031.260, 2033.250",
        }

    def get_motion_to_compel_deadline(
        self,
        response_date: datetime
    ) -> Dict[str, Any]:
        """
        Calculate motion to compel further responses deadline (45 calendar days).

        Args:
            response_date: Date discovery response was served

        Returns:
            Deadline dict with 'title', 'date', 'description'
        """
        # 45 calendar days from response
        deadline = self.count_calendar_days(response_date, 45, 'forward')

        # Apply CCP 12a
        final_deadline = self.apply_ccp_12a_adjustment(deadline, 'forward')

        return {
            'title': 'Motion to Compel Further Due',
            'date': final_deadline,
            'description': "Motion to compel further responses due 45 calendar days from service of response\nCCP § 2030.300(c), 2031.310(c), 2033.290(c)",
        }

    def get_opposition_deadline(
        self,
        motion_type: str,
        hearing_date: datetime,
        service_method: str = 'electronic'
    ) -> Optional[Dict[str, Any]]:
        """
        Get opposition deadline for a motion.

        Args:
            motion_type: Motion type ('msj', 'standard', 'demurrer', etc.)
            hearing_date: The hearing date
            service_method: Service method

        Returns:
            Deadline dict or None
        """
        motion_prefixes = {
            'msj': 'msj',
            'msa': 'msj',  # MSA uses same deadlines as MSJ
            'standard': 'standard-motion',
            'demurrer': 'demurrer',
            'discovery': 'discovery-motion',
        }

        prefix = motion_prefixes.get(motion_type.lower(), 'standard-motion')
        rule_slug = f"{prefix}-opposition"

        rule = self.rules_by_slug.get(rule_slug)
        if not rule:
            return None

        deadline_date = self.calculate_deadline(
            rule_slug=rule_slug,
            trigger_date=hearing_date,
            service_method=service_method
        )

        if deadline_date:
            return {
                'title': 'Opposition Due',
                'date': deadline_date,
                'description': f"{rule.get('name', '')}\nStatute: {rule.get('statute', '')}\nHearing: {hearing_date.strftime('%Y-%m-%d')}\nService: {service_method}",
                'rule_slug': rule_slug,
            }

        return None

    def get_reply_deadline(
        self,
        motion_type: str,
        hearing_date: datetime,
        service_method: str = 'electronic'
    ) -> Optional[Dict[str, Any]]:
        """
        Get reply deadline for a motion.

        Args:
            motion_type: Motion type ('msj', 'standard', 'demurrer', etc.)
            hearing_date: The hearing date
            service_method: Service method

        Returns:
            Deadline dict or None
        """
        motion_prefixes = {
            'msj': 'msj',
            'msa': 'msj',
            'standard': 'standard-motion',
            'demurrer': 'demurrer',
            'discovery': 'discovery-motion',
        }

        prefix = motion_prefixes.get(motion_type.lower(), 'standard-motion')
        rule_slug = f"{prefix}-reply"

        rule = self.rules_by_slug.get(rule_slug)
        if not rule:
            return None

        deadline_date = self.calculate_deadline(
            rule_slug=rule_slug,
            trigger_date=hearing_date,
            service_method=service_method
        )

        if deadline_date:
            return {
                'title': 'Reply Due',
                'date': deadline_date,
                'description': f"{rule.get('name', '')}\nStatute: {rule.get('statute', '')}\nHearing: {hearing_date.strftime('%Y-%m-%d')}\nService: {service_method}",
                'rule_slug': rule_slug,
            }

        return None
