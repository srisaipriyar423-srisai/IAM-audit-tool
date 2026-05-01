"""
Segregation of Duties (SoD) analysis engine.
Aggregates SoD violations across all analysed users into a DataFrame.
"""

import pandas as pd


def get_all_sod_violations(analyzed_users: list) -> pd.DataFrame:
    """
    Parameters
    ----------
    analyzed_users : list[dict]
        Each dict is a user record merged with its risk-engine output.

    Returns
    -------
    pd.DataFrame
        Columns: User, Department, Job Title, Conflict, Severity, Risk Impact
    """
    rows = []
    for u in analyzed_users:
        for conflict in u.get("sod_conflicts", []):
            severity = conflict["severity"]
            if severity == "Critical":
                impact = "+20 risk points"
            elif severity == "High":
                impact = "+12 risk points"
            else:
                impact = "+6 risk points"

            rows.append(
                {
                    "User": u["name"],
                    "Department": u["department"],
                    "Job Title": u["job_title"],
                    "Conflict": conflict["conflict_name"],
                    "Severity": severity,
                    "Risk Impact": impact,
                }
            )

    if not rows:
        return pd.DataFrame(
            columns=["User", "Department", "Job Title", "Conflict", "Severity", "Risk Impact"]
        )

    df = pd.DataFrame(rows)
    severity_order = {"Critical": 0, "High": 1, "Medium": 2}
    df["_sort"] = df["Severity"].map(severity_order)
    df = df.sort_values("_sort").drop(columns="_sort").reset_index(drop=True)
    return df
