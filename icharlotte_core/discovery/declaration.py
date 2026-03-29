"""
Declaration generator for discovery sets exceeding CCP limits.

When Special Interrogatories or Requests for Admission exceed the 35-request
statutory limit, California Code of Civil Procedure requires a supporting
declaration of necessity.  This module generates the declaration text.
"""

from .models import DiscoverySet, DiscoveryType, number_to_word


# CCP sections keyed by discovery type
_CCP_LIMIT_SECTION = {
    DiscoveryType.SI: "2030.030",
    DiscoveryType.RFA: "2033.030",
}

_CCP_DECLARATION_SECTION = {
    DiscoveryType.SI: "2030.070",
    DiscoveryType.RFA: "2033.050",
}

_TYPE_LONG_NAME = {
    DiscoveryType.SI: "Special Interrogatories",
    DiscoveryType.RFA: "Requests for Admission",
}

_TYPE_TITLE_NAME = {
    DiscoveryType.SI: "SPECIAL INTERROGATORIES",
    DiscoveryType.RFA: "REQUESTS FOR ADMISSION",
}


def generate_declaration(
    discovery_set: DiscoverySet,
    attorney_name: str,
    firm_name: str = "Bordin Semmer LLP",
) -> str:
    """Generate a CCP declaration of necessity for exceeding the 35-request limit.

    Returns an empty string when no declaration is required:
    - RPDs never need declarations
    - Sets with total_count <= 35 are within the statutory limit

    Args:
        discovery_set: The discovery set that may require a declaration.
        attorney_name: Name of the attorney signing the declaration.
        firm_name: Name of the law firm (default: Bordin Semmer LLP).

    Returns:
        Declaration text with a ``{date}`` placeholder for the execution date,
        or an empty string if no declaration is needed.
    """
    ds = discovery_set

    # RPDs never need declarations
    if ds.discovery_type == DiscoveryType.RPD:
        return ""

    # Under or at the 35-request limit — no declaration needed
    if ds.total_count <= 35:
        return ""

    dtype = ds.discovery_type
    limit_section = _CCP_LIMIT_SECTION[dtype]
    decl_section = _CCP_DECLARATION_SECTION[dtype]
    type_long = _TYPE_LONG_NAME[dtype]
    type_title = _TYPE_TITLE_NAME[dtype]

    current_count = len(ds.requests)
    previous_count = ds.previous_count
    total_count = ds.total_count

    # Word forms for counts
    if previous_count == 0:
        prev_word_str = "zero (0)"
    else:
        try:
            prev_word = number_to_word(previous_count)
            prev_word_str = f"{prev_word.lower()} ({previous_count})"
        except ValueError:
            prev_word_str = f"{previous_count}"

    try:
        current_word = number_to_word(current_count)
        current_word_str = f"{current_word.lower()} ({current_count})"
    except ValueError:
        current_word_str = f"{current_count}"

    try:
        total_word = number_to_word(total_count)
        total_word_str = f"{total_word.lower()} ({total_count})"
    except ValueError:
        total_word_str = f"{total_count}"

    set_word = ds.set_word
    propounding = ds.propounding_party.name
    directed_to = ds.directed_to.name
    directed_role = ds.directed_to.role_label

    title = (
        f"DECLARATION OF {attorney_name.upper()} IN SUPPORT OF "
        f"ADDITIONAL {type_title} PROPOUNDED ON {directed_to.upper()}"
    )

    lines = [
        title,
        "",
        (
            f"I, {attorney_name}, declare as follows:"
        ),
        "",
        (
            f"1. I am an attorney at law duly licensed to practice before all courts "
            f"of the State of California. I am a member of {firm_name}, attorneys of "
            f"record for {propounding} in this action."
        ),
        "",
        (
            f"2. I am propounding on behalf of {propounding} the attached "
            f"{type_long}, Set {set_word}, directed to {directed_role}, "
            f"{directed_to}."
        ),
        "",
        (
            f"3. This set of {type_long.lower()} will cause the total number of "
            f"{type_long.lower()} propounded to {directed_to} to exceed the limit "
            f"of thirty-five (35) permitted under Code of Civil Procedure section "
            f"{limit_section}. This declaration is made pursuant to Code of Civil "
            f"Procedure section {decl_section}."
        ),
        "",
        (
            f"4. {propounding} has previously propounded {prev_word_str} "
            f"{type_long.lower()} on {directed_to}. This set contains "
            f"{current_word_str} {type_long.lower()}, bringing the total to "
            f"{total_word_str}."
        ),
        "",
        (
            f"5. I have personally examined each of the {type_long.lower()} in this "
            f"set. The number of {type_long.lower()} propounded is warranted by the "
            f"complexity and nature of the issues in this litigation."
        ),
        "",
        (
            f"6. None of the {type_long.lower()} in this set has been propounded "
            f"for any improper purpose, such as to harass the responding party or to "
            f"cause unnecessary delay or needless increase in the cost of litigation."
        ),
        "",
        (
            f"I declare under penalty of perjury under the laws of the State of "
            f"California that the foregoing is true and correct."
        ),
        "",
        f"Executed on {{date}}.",
        "",
        "",
        f"____________________________",
        f"{attorney_name}",
    ]

    return "\n".join(lines)
