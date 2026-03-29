# Discovery module for iCharlotte
# Provides discovery generation pipeline: models, templates, set tracking, assembly

from .models import (
    PartyRole,
    DiscoveryMode,
    CustomStyle,
    DiscoveryType,
    Party,
    DiscoveryRequest,
    DiscoverySet,
    SetTrackerResult,
    number_to_word,
    generate_abbreviation,
)

from .templates import (
    TemplateLoader,
    extract_requests_from_text,
    substitute_variables,
)

from .set_tracker import SetTracker
from .declaration import generate_declaration

__all__ = [
    'PartyRole',
    'DiscoveryMode',
    'CustomStyle',
    'DiscoveryType',
    'Party',
    'DiscoveryRequest',
    'DiscoverySet',
    'SetTrackerResult',
    'number_to_word',
    'generate_abbreviation',
    'TemplateLoader',
    'extract_requests_from_text',
    'substitute_variables',
    'SetTracker',
    'generate_declaration',
]
