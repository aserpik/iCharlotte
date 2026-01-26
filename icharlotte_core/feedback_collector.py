"""
Feedback Collector for iCharlotte

Collects and stores user feedback on agent outputs for quality improvement.

Features:
- Rating collection (1-5 stars)
- Issue categorization
- Free-text corrections
- Feedback storage and retrieval
- Analytics support
"""

import os
import json
import uuid
from datetime import datetime
from typing import List, Optional, Dict, Any
from dataclasses import dataclass, field, asdict


# =============================================================================
# Configuration
# =============================================================================

FEEDBACK_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), ".gemini", "feedback")


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class FeedbackIssue:
    """A specific issue reported in feedback."""
    category: str  # missing_info, incorrect, formatting, other
    description: str = ""
    severity: str = "medium"  # low, medium, high


@dataclass
class Feedback:
    """User feedback for an agent output."""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())

    # Context
    agent_name: str = ""
    file_number: str = ""
    document_name: str = ""
    prompt_version: str = ""

    # Rating
    rating: int = 0  # 1-5 stars, 0 = no rating

    # Issues
    issues: List[FeedbackIssue] = field(default_factory=list)

    # Corrections
    correction_text: str = ""
    original_text: str = ""

    # Additional notes
    notes: str = ""

    def to_dict(self) -> dict:
        result = asdict(self)
        result['issues'] = [asdict(i) for i in self.issues]
        return result

    @classmethod
    def from_dict(cls, data: dict) -> 'Feedback':
        issues = [FeedbackIssue(**i) for i in data.pop('issues', [])]
        return cls(**data, issues=issues)


# =============================================================================
# Feedback Collector
# =============================================================================

class FeedbackCollector:
    """
    Collects and stores user feedback.

    Usage:
        collector = FeedbackCollector()

        # Create feedback
        feedback = Feedback(
            agent_name="Discovery",
            file_number="2024.123",
            rating=4,
            issues=[FeedbackIssue(category="missing_info", description="Missing treatment dates")]
        )

        # Save feedback
        collector.save_feedback(feedback)

        # Get feedback for analysis
        all_feedback = collector.get_feedback_for_agent("Discovery")
    """

    ISSUE_CATEGORIES = [
        ("missing_info", "Missing Information"),
        ("incorrect", "Incorrect Information"),
        ("formatting", "Formatting Issues"),
        ("incomplete", "Incomplete Summary"),
        ("redundant", "Redundant Content"),
        ("unclear", "Unclear Writing"),
        ("other", "Other Issue"),
    ]

    def __init__(self, feedback_dir: str = None):
        """
        Initialize the feedback collector.

        Args:
            feedback_dir: Directory for feedback storage.
        """
        self.feedback_dir = feedback_dir or FEEDBACK_DIR
        self._ensure_directory()

    def _ensure_directory(self):
        """Create feedback directory if needed."""
        if not os.path.exists(self.feedback_dir):
            os.makedirs(self.feedback_dir)

    def _get_file_path(self, agent_name: str, date_str: str = None) -> str:
        """Get the file path for storing feedback."""
        if date_str is None:
            date_str = datetime.now().strftime("%Y-%m-%d")
        return os.path.join(self.feedback_dir, f"{agent_name}_{date_str}.json")

    def save_feedback(self, feedback: Feedback) -> bool:
        """
        Save feedback to storage.

        Args:
            feedback: Feedback object to save.

        Returns:
            True if successful.
        """
        try:
            file_path = self._get_file_path(feedback.agent_name)

            # Load existing feedback for this file
            existing = []
            if os.path.exists(file_path):
                with open(file_path, "r", encoding="utf-8") as f:
                    existing = json.load(f)

            # Append new feedback
            existing.append(feedback.to_dict())

            # Save
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(existing, f, indent=2)

            return True

        except Exception as e:
            print(f"Error saving feedback: {e}")
            return False

    def get_feedback_for_agent(
        self,
        agent_name: str,
        start_date: str = None,
        end_date: str = None
    ) -> List[Feedback]:
        """
        Get all feedback for an agent.

        Args:
            agent_name: Name of the agent.
            start_date: Optional start date filter (YYYY-MM-DD).
            end_date: Optional end date filter (YYYY-MM-DD).

        Returns:
            List of Feedback objects.
        """
        feedback_list = []

        # Find all feedback files for this agent
        for filename in os.listdir(self.feedback_dir):
            if filename.startswith(f"{agent_name}_") and filename.endswith(".json"):
                # Extract date from filename
                date_str = filename[len(agent_name)+1:-5]

                # Apply date filters
                if start_date and date_str < start_date:
                    continue
                if end_date and date_str > end_date:
                    continue

                file_path = os.path.join(self.feedback_dir, filename)
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        for item in data:
                            feedback_list.append(Feedback.from_dict(item))
                except Exception:
                    continue

        return feedback_list

    def get_feedback_for_case(self, file_number: str) -> List[Feedback]:
        """
        Get all feedback for a specific case.

        Args:
            file_number: Case file number.

        Returns:
            List of Feedback objects.
        """
        feedback_list = []

        for filename in os.listdir(self.feedback_dir):
            if filename.endswith(".json"):
                file_path = os.path.join(self.feedback_dir, filename)
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        for item in data:
                            if item.get("file_number") == file_number:
                                feedback_list.append(Feedback.from_dict(item))
                except Exception:
                    continue

        return feedback_list

    def get_statistics(self, agent_name: str = None) -> Dict[str, Any]:
        """
        Get statistics about collected feedback.

        Args:
            agent_name: Optional filter by agent.

        Returns:
            Dictionary with statistics.
        """
        if agent_name:
            feedback_list = self.get_feedback_for_agent(agent_name)
        else:
            feedback_list = []
            for filename in os.listdir(self.feedback_dir):
                if filename.endswith(".json"):
                    file_path = os.path.join(self.feedback_dir, filename)
                    try:
                        with open(file_path, "r", encoding="utf-8") as f:
                            data = json.load(f)
                            for item in data:
                                feedback_list.append(Feedback.from_dict(item))
                    except Exception:
                        continue

        if not feedback_list:
            return {
                'total_feedback': 0,
                'average_rating': 0,
                'rating_distribution': {},
                'common_issues': {},
            }

        # Calculate statistics
        total = len(feedback_list)
        ratings = [f.rating for f in feedback_list if f.rating > 0]
        avg_rating = sum(ratings) / len(ratings) if ratings else 0

        # Rating distribution
        rating_dist = {i: 0 for i in range(1, 6)}
        for r in ratings:
            rating_dist[r] += 1

        # Issue frequency
        issue_freq = {}
        for feedback in feedback_list:
            for issue in feedback.issues:
                cat = issue.category
                issue_freq[cat] = issue_freq.get(cat, 0) + 1

        return {
            'total_feedback': total,
            'total_rated': len(ratings),
            'average_rating': round(avg_rating, 2),
            'rating_distribution': rating_dist,
            'common_issues': issue_freq,
        }

    def get_low_rated_outputs(
        self,
        agent_name: str,
        threshold: int = 3
    ) -> List[Feedback]:
        """
        Get feedback for outputs rated below threshold.

        Args:
            agent_name: Agent name.
            threshold: Rating threshold (inclusive).

        Returns:
            List of low-rated feedback.
        """
        all_feedback = self.get_feedback_for_agent(agent_name)
        return [f for f in all_feedback if 0 < f.rating <= threshold]

    def export_for_training(
        self,
        agent_name: str,
        min_rating: int = 4
    ) -> List[Dict]:
        """
        Export high-quality feedback for training/fine-tuning.

        Args:
            agent_name: Agent name.
            min_rating: Minimum rating to include.

        Returns:
            List of training examples.
        """
        all_feedback = self.get_feedback_for_agent(agent_name)
        training_data = []

        for feedback in all_feedback:
            if feedback.rating >= min_rating and feedback.correction_text:
                training_data.append({
                    'original': feedback.original_text,
                    'corrected': feedback.correction_text,
                    'rating': feedback.rating,
                    'document': feedback.document_name,
                })

        return training_data


# =============================================================================
# UI Dialog Data
# =============================================================================

def get_issue_categories() -> List[tuple]:
    """Get list of issue categories for UI."""
    return FeedbackCollector.ISSUE_CATEGORIES


def create_feedback(
    agent_name: str,
    file_number: str = "",
    document_name: str = "",
    rating: int = 0,
    issues: List[tuple] = None,  # [(category, description), ...]
    correction: str = "",
    original: str = "",
    notes: str = ""
) -> Feedback:
    """
    Convenience function to create feedback.

    Args:
        agent_name: Name of the agent.
        file_number: Case file number.
        document_name: Name of the processed document.
        rating: Rating 1-5.
        issues: List of (category, description) tuples.
        correction: Corrected text if user provided one.
        original: Original output text.
        notes: Additional notes.

    Returns:
        Feedback object.
    """
    feedback_issues = []
    if issues:
        for cat, desc in issues:
            feedback_issues.append(FeedbackIssue(category=cat, description=desc))

    return Feedback(
        agent_name=agent_name,
        file_number=file_number,
        document_name=document_name,
        rating=rating,
        issues=feedback_issues,
        correction_text=correction,
        original_text=original,
        notes=notes
    )


# =============================================================================
# Global Instance
# =============================================================================

_default_collector: Optional[FeedbackCollector] = None


def get_collector() -> FeedbackCollector:
    """Get the global feedback collector instance."""
    global _default_collector
    if _default_collector is None:
        _default_collector = FeedbackCollector()
    return _default_collector


def save_feedback(feedback: Feedback) -> bool:
    """Convenience function to save feedback."""
    return get_collector().save_feedback(feedback)


def get_statistics(agent_name: str = None) -> Dict[str, Any]:
    """Convenience function to get statistics."""
    return get_collector().get_statistics(agent_name)
