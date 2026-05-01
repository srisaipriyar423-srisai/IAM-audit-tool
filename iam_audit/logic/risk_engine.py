"""
Risk scoring engine for IAM User Access Review.
Calculates per-user risk scores based on access, SoD, MFA, password age, etc.
"""

from datetime import datetime, date

REFERENCE_DATE = date(2025, 7, 1)


def calculate_risk(
    user: dict,
    rbac_matrix: dict,
    sod_rules: list,
    role_catalog: dict,
) -> dict:
    """Analyse a single user and return a comprehensive risk assessment dict."""

    assigned = set(user["assigned_roles"])
    expected = set(rbac_matrix.get(user["job_title"], []))

    excess_roles = sorted(assigned - expected)
    missing_roles = sorted(expected - assigned)

    # ── SoD conflicts ────────────────────────────────────────────────
    sod_conflicts = []
    for rule in sod_rules:
        if rule["role_a"] in assigned and rule["role_b"] in assigned:
            sod_conflicts.append(rule)

    # ── Critical roles held ──────────────────────────────────────────
    critical_roles = [
        r for r in user["assigned_roles"]
        if role_catalog.get(r, {}).get("sensitivity") == "Critical"
    ]

    # ── Days since last login ────────────────────────────────────────
    last_login = datetime.strptime(user["last_login"], "%Y-%m-%d").date()
    days_since_login = (REFERENCE_DATE - last_login).days

    # ── Risk score calculation (0-100) ───────────────────────────────
    score = 0
    risk_flags: list[str] = []

    # Employment status
    if user["employment_status"] == "Terminated" and assigned:
        score += 30
        risk_flags.append("Terminated employee still has active access roles")
    # Note: James Ford has department "Terminated" but status "Active" – treat
    # department name as an anomaly flag too.
    if user["department"] == "Terminated" and assigned:
        if "Terminated employee still has active access roles" not in risk_flags:
            score += 30
            risk_flags.append("Terminated employee still has active access roles")

    if user["employment_status"] == "Inactive" and assigned:
        score += 10
        risk_flags.append("Inactive employee retains access roles")

    # MFA
    if not user["mfa_enabled"]:
        score += 15
        risk_flags.append("Multi-Factor Authentication (MFA) is disabled")

    # Password age
    if user["password_age_days"] > 90:
        score += 10
        risk_flags.append(
            f"Password age exceeds 90-day policy ({user['password_age_days']} days)"
        )

    # Excess roles (capped at +20)
    if excess_roles:
        excess_pts = min(len(excess_roles) * 5, 20)
        score += excess_pts
        risk_flags.append(
            f"User has {len(excess_roles)} excess role(s): {', '.join(excess_roles)}"
        )

    # SoD conflicts
    for conflict in sod_conflicts:
        sev = conflict["severity"]
        if sev == "Critical":
            score += 20
        elif sev == "High":
            score += 12
        else:
            score += 6
        risk_flags.append(
            f"SoD conflict ({sev}): {conflict['conflict_name']}"
        )

    # Login dormancy
    if days_since_login > 90:
        score += 10
        risk_flags.append(
            f"Account dormant — last login was {days_since_login} days ago"
        )

    # Critical role concentration
    if len(critical_roles) > 2:
        score += 5
        risk_flags.append(
            f"Holds {len(critical_roles)} critical-sensitivity roles"
        )

    # Cap at 100
    score = min(score, 100)

    # ── Risk level & compliance ──────────────────────────────────────
    if score >= 40:
        risk_level = "Critical"
    elif score >= 25:
        risk_level = "High"
    elif score >= 12:
        risk_level = "Medium"
    else:
        risk_level = "Low"

    if risk_level in ("Critical", "High"):
        compliance_status = "Non-Compliant"
    elif risk_level == "Medium":
        compliance_status = "Review Required"
    else:
        compliance_status = "Compliant"

    return {
        "excess_roles": excess_roles,
        "missing_roles": missing_roles,
        "sod_conflicts": sod_conflicts,
        "critical_roles": critical_roles,
        "risk_score": score,
        "risk_level": risk_level,
        "risk_flags": risk_flags,
        "days_since_login": days_since_login,
        "compliance_status": compliance_status,
    }
