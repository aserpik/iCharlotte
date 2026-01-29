"""
Document Registry for iCharlotte

Tracks all summarized documents for each case with their classifications.
Used by contradiction detection and timeline agents to allow pre-selection of documents.
"""

import os
import json
import datetime
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict

# Add parent directory to path for imports
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from icharlotte_core.llm_config import LLMCaller


# =============================================================================
# Document Types
# =============================================================================

DOCUMENT_TYPES = [
    # Legal Pleadings
    "Complaint",
    "Answer",
    "Cross-Complaint",
    "Demurrer",
    "Motion",
    "Opposition",
    "Reply",
    "Court Order",
    "Judgment",

    # Discovery
    "Form Interrogatories - Responses",
    "Special Interrogatories - Responses",
    "Request for Admissions - Responses",
    "Request for Production - Responses",
    "Form Interrogatories - Propounded",
    "Special Interrogatories - Propounded",
    "Request for Admissions - Propounded",
    "Request for Production - Propounded",

    # Depositions
    "Deposition - Plaintiff",
    "Deposition - Defendant",
    "Deposition - Witness",
    "Deposition - Expert",
    "Deposition - Corporate Representative",

    # Reports
    "Traffic Collision Report",
    "Police Report",
    "Incident Report",
    "ISO ClaimSearch Report",
    "Investigation Report",
    "Expert Report",
    "IME Report",
    "Biomechanical Report",

    # Medical Records
    "Medical Records - Hospital",
    "Medical Records - Emergency Room",
    "Medical Records - Primary Care",
    "Medical Records - Specialist",
    "Medical Records - Imaging",
    "Medical Records - Physical Therapy",
    "Medical Records - Chiropractic",
    "Medical Bills",
    "Medical Chronology",

    # Contracts/Agreements
    "Lease Agreement",
    "Contract",
    "Insurance Policy",
    "Settlement Agreement",
    "Release",

    # Correspondence
    "Demand Letter",
    "Attorney Correspondence",
    "Insurance Correspondence",
    "Client Correspondence",

    # Other
    "Photographs",
    "Video Evidence",
    "Witness Statement",
    "Declaration",
    "Affidavit",
    "Subpoena",
    "Other",
]


# =============================================================================
# Classification Prompt
# =============================================================================

CLASSIFICATION_PROMPT = """You are a legal document classifier. Based on the document summary provided, classify the document into ONE of the following categories.

CATEGORIES:
{categories}

INSTRUCTIONS:
1. Read the summary carefully
2. Identify the document type based on its content and structure
3. Return ONLY the category name, exactly as written above
4. If the document doesn't fit any category well, return "Other"

For depositions, identify WHO was deposed (plaintiff, defendant, witness, expert, or corporate representative).
For discovery responses, identify the TYPE (form interrogatories, special interrogatories, requests for admission, or requests for production).
For medical records, identify the FACILITY TYPE if possible.

SUMMARY:
{summary}

DOCUMENT TYPE:"""


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class RegisteredDocument:
    """A document that has been summarized and registered."""
    name: str                    # Original filename
    document_type: str           # Classified type
    source_path: str             # Path to original document
    summary_location: str        # Where summary is stored (docx path or case data key)
    agent: str                   # Which agent processed it (summarize, summarize_discovery, etc.)
    timestamp: str               # When it was processed
    char_count: int = 0          # Character count of summary

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'RegisteredDocument':
        return cls(**data)


# =============================================================================
# Document Registry
# =============================================================================

class DocumentRegistry:
    """
    Manages the registry of summarized documents per case.

    Storage: .gemini/case_data/{file_number}_document_registry.json
    """

    def __init__(self, base_dir: str = None):
        if base_dir:
            self.data_dir = base_dir
        else:
            self.data_dir = os.path.join(os.getcwd(), ".gemini", "case_data")

        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

    def _get_registry_path(self, file_number: str) -> str:
        return os.path.join(self.data_dir, f"{file_number}_document_registry.json")

    def _load_registry(self, file_number: str) -> List[dict]:
        path = self._get_registry_path(file_number)
        if not os.path.exists(path):
            return []
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []

    def _save_registry(self, file_number: str, documents: List[dict]):
        path = self._get_registry_path(file_number)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(documents, f, indent=2)

    def register_document(
        self,
        file_number: str,
        name: str,
        document_type: str,
        source_path: str,
        summary_location: str,
        agent: str,
        char_count: int = 0
    ) -> RegisteredDocument:
        """
        Register a newly summarized document.

        Updates existing entry if document with same name exists.
        """
        documents = self._load_registry(file_number)

        doc = RegisteredDocument(
            name=name,
            document_type=document_type,
            source_path=source_path,
            summary_location=summary_location,
            agent=agent,
            timestamp=datetime.datetime.now().isoformat(),
            char_count=char_count
        )

        # Check for existing entry with same name
        for i, existing in enumerate(documents):
            if existing.get('name') == name:
                documents[i] = doc.to_dict()
                break
        else:
            documents.append(doc.to_dict())

        self._save_registry(file_number, documents)
        return doc

    def get_all_documents(self, file_number: str) -> List[RegisteredDocument]:
        """Get all registered documents for a case."""
        documents = self._load_registry(file_number)
        return [RegisteredDocument.from_dict(d) for d in documents]

    def get_documents_by_type(self, file_number: str, doc_types: List[str]) -> List[RegisteredDocument]:
        """Get documents matching specific types."""
        all_docs = self.get_all_documents(file_number)
        return [d for d in all_docs if d.document_type in doc_types]

    def get_document_types(self, file_number: str) -> List[str]:
        """Get list of unique document types for a case."""
        all_docs = self.get_all_documents(file_number)
        return list(set(d.document_type for d in all_docs))

    def get_documents_by_agent(self, file_number: str, agent: str) -> List[RegisteredDocument]:
        """Get documents processed by a specific agent."""
        all_docs = self.get_all_documents(file_number)
        return [d for d in all_docs if d.agent == agent]

    def remove_document(self, file_number: str, name: str) -> bool:
        """Remove a document from the registry."""
        documents = self._load_registry(file_number)
        original_len = len(documents)
        documents = [d for d in documents if d.get('name') != name]

        if len(documents) < original_len:
            self._save_registry(file_number, documents)
            return True
        return False


# =============================================================================
# Document Classifier
# =============================================================================

class DocumentClassifier:
    """
    Classifies documents based on their summary content.
    """

    def __init__(self, logger=None):
        self.logger = logger
        self.llm_caller = LLMCaller(logger=logger)

    def classify(self, summary: str, filename: str = "") -> str:
        """
        Classify a document based on its summary.

        Args:
            summary: The document summary text.
            filename: Original filename (can help with classification).

        Returns:
            Document type string.
        """
        # Quick classification based on filename patterns
        filename_lower = filename.lower()

        # Check for obvious patterns first
        if "depo" in filename_lower or "deposition" in filename_lower:
            return self._classify_deposition(summary)

        if "interrog" in filename_lower or "frog" in filename_lower or "srog" in filename_lower:
            return self._classify_discovery(summary, filename_lower)

        if "rfa" in filename_lower or "admission" in filename_lower:
            return self._classify_discovery(summary, filename_lower)

        if "rfp" in filename_lower or "production" in filename_lower:
            return self._classify_discovery(summary, filename_lower)

        if "complaint" in filename_lower:
            if "cross" in filename_lower:
                return "Cross-Complaint"
            return "Complaint"

        if "answer" in filename_lower:
            return "Answer"

        if "traffic collision" in filename_lower or "tcr" in filename_lower:
            return "Traffic Collision Report"

        if "police report" in filename_lower:
            return "Police Report"

        if "iso" in filename_lower or "claimsearch" in filename_lower:
            return "ISO ClaimSearch Report"

        # Use LLM for classification
        return self._llm_classify(summary)

    def _classify_deposition(self, summary: str) -> str:
        """Classify deposition subtype."""
        summary_lower = summary.lower()

        if "plaintiff" in summary_lower[:2000]:
            return "Deposition - Plaintiff"
        if "defendant" in summary_lower[:2000]:
            return "Deposition - Defendant"
        if "expert" in summary_lower[:2000] or "opinion" in summary_lower[:1000]:
            return "Deposition - Expert"
        if "pmo" in summary_lower or "person most qualified" in summary_lower or "corporate" in summary_lower:
            return "Deposition - Corporate Representative"

        return "Deposition - Witness"

    def _classify_discovery(self, summary: str, filename_lower: str) -> str:
        """Classify discovery subtype."""
        is_response = "response" in filename_lower or "resp" in filename_lower

        if "frog" in filename_lower or "form interrog" in filename_lower:
            return "Form Interrogatories - Responses" if is_response else "Form Interrogatories - Propounded"

        if "srog" in filename_lower or "special interrog" in filename_lower:
            return "Special Interrogatories - Responses" if is_response else "Special Interrogatories - Propounded"

        if "rfa" in filename_lower or "admission" in filename_lower:
            return "Request for Admissions - Responses" if is_response else "Request for Admissions - Propounded"

        if "rfp" in filename_lower or "production" in filename_lower:
            return "Request for Production - Responses" if is_response else "Request for Production - Propounded"

        # Default based on response status
        return "Form Interrogatories - Responses" if is_response else "Form Interrogatories - Propounded"

    def _llm_classify(self, summary: str) -> str:
        """Use LLM to classify document."""
        try:
            categories = "\n".join(f"- {t}" for t in DOCUMENT_TYPES)
            prompt = CLASSIFICATION_PROMPT.format(
                categories=categories,
                summary=summary[:5000]  # Limit summary length
            )

            response = self.llm_caller.call(
                prompt,
                "",
                agent_id="agent_chat"  # Use fast model for classification
            )

            if response:
                # Clean up response
                response = response.strip()

                # Find matching document type
                for doc_type in DOCUMENT_TYPES:
                    if doc_type.lower() in response.lower():
                        return doc_type

                # If exact match not found, return cleaned response or Other
                if response in DOCUMENT_TYPES:
                    return response

            return "Other"

        except Exception as e:
            if self.logger:
                self.logger.warning(f"LLM classification failed: {e}")
            return "Other"


# =============================================================================
# Utility Functions
# =============================================================================

def classify_and_register(
    file_number: str,
    document_name: str,
    summary: str,
    source_path: str,
    summary_location: str,
    agent: str,
    logger=None
) -> RegisteredDocument:
    """
    Convenience function to classify a document and register it.

    Args:
        file_number: Case file number.
        document_name: Original document filename.
        summary: The summary text.
        source_path: Path to original document.
        summary_location: Where summary is stored.
        agent: Which agent processed it.
        logger: Optional logger.

    Returns:
        RegisteredDocument instance.
    """
    classifier = DocumentClassifier(logger=logger)
    doc_type = classifier.classify(summary, document_name)

    if logger:
        logger.info(f"Classified '{document_name}' as: {doc_type}")

    registry = DocumentRegistry()
    doc = registry.register_document(
        file_number=file_number,
        name=document_name,
        document_type=doc_type,
        source_path=source_path,
        summary_location=summary_location,
        agent=agent,
        char_count=len(summary)
    )

    return doc


def get_available_documents(file_number: str) -> List[Dict[str, Any]]:
    """
    Get list of available documents for a case (for UI display).

    Returns list of dicts with: name, document_type, agent, timestamp
    """
    registry = DocumentRegistry()
    docs = registry.get_all_documents(file_number)

    return [
        {
            "name": d.name,
            "document_type": d.document_type,
            "agent": d.agent,
            "timestamp": d.timestamp,
            "char_count": d.char_count
        }
        for d in docs
    ]


def get_document_type_list() -> List[str]:
    """Get the full list of document types for UI dropdowns."""
    return DOCUMENT_TYPES.copy()
