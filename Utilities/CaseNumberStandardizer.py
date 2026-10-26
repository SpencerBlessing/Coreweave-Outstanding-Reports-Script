import pandas as pd


def standardize_case_number(value):
    """Prefix numeric case IDs with CS- and pad to 8 digits."""
    if pd.isna(value):
        return value

    value = str(value).strip().upper()

    if value.startswith("CS-"):
        number = value[3:]
        if number.isdigit():
            return f"CS-{number.zfill(8)}"
        return value

    if value.isdigit():
        return f"CS-{value.zfill(8)}"

    return value