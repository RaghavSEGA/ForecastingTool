"""Sign-in gate, same scheme as the other SEGA Streamlit tools.

@segaamerica.com email -> 6-digit one-time code sent via AWS SES (10-minute expiry,
5 attempts) -> HMAC-signed ?t= URL token valid for 1 day.

Secrets (.streamlit/secrets.toml):
    AWS_SES_REGION, AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, EMAIL_FROM, COOKIE_SIGNING_KEY

Usage, straight after st.set_page_config():
    from auth import require_login
    require_login()
"""
from __future__ import annotations

import hashlib
import hmac
import secrets
import time

import boto3
import streamlit as st

ALLOWED_DOMAIN = "@segaamerica.com"
OTP_EXPIRY = 600          # 10 minutes
TOKEN_EXPIRY = 86400      # 1 day
MAX_ATTEMPTS = 5
APP_NAME = "SEGA Forecast Consolidator"


# ------------------------------------------------------------------ tokens --
def _key() -> bytes | None:
    try:
        key = st.secrets.get("COOKIE_SIGNING_KEY")
    except Exception:  # no secrets.toml at all
        key = None
    return key.encode() if key else None


def _sign(payload: str) -> str:
    return hmac.new(_key(), payload.encode(), hashlib.sha256).hexdigest()


def _make_token(email: str) -> str:
    payload = f"{email}:{int(time.time()) + TOKEN_EXPIRY}"
    return f"{payload}:{_sign(payload)}"


def _verify_token(token: str) -> str | None:
    parts = token.split(":")
    if len(parts) != 3:
        return None
    email, expiry, sig = parts
    try:
        if time.time() > int(expiry):
            return None
    except ValueError:
        return None
    if not hmac.compare_digest(sig, _sign(f"{email}:{expiry}")):
        return None
    return email if email.endswith(ALLOWED_DOMAIN) else None


# --------------------------------------------------------------------- SES --
def _send_otp(email: str, otp: str) -> bool:
    try:
        ses = boto3.client(
            "ses",
            region_name=st.secrets.get("AWS_SES_REGION", "us-east-1"),
            aws_access_key_id=st.secrets.get("AWS_ACCESS_KEY_ID"),
            aws_secret_access_key=st.secrets.get("AWS_SECRET_ACCESS_KEY"),
        )
        ses.send_email(
            Source=st.secrets.get("EMAIL_FROM", "noreply@segaamerica.com"),
            Destination={"ToAddresses": [email]},
            Message={
                "Subject": {"Data": f"{APP_NAME} — Your sign-in code"},
                "Body": {
                    "Text": {"Data": f"Your one-time code is: {otp}\n\n"
                                     "Expires in 10 minutes. Do not share this code."},
                    "Html": {"Data": f"""<!DOCTYPE html>
<html><body style="font-family:Arial,sans-serif;background:#f4f6f9;color:#222;padding:40px;margin:0">
<div style="max-width:480px;margin:auto;background:#fff;border-radius:8px;padding:32px;
            border:1px solid #d9dee7">
  <h2 style="color:#1F3864;margin-top:0;font-size:1.3rem">📊 {APP_NAME}</h2>
  <p>Your one-time sign-in code:</p>
  <div style="font-size:36px;font-weight:700;letter-spacing:10px;color:#1F3864;
              background:#eef2f8;padding:18px;border-radius:6px;text-align:center;
              margin:16px 0">{otp}</div>
  <p style="color:#777;font-size:12px;margin-top:24px">
    Expires in 10 minutes. Do not share this code.</p>
</div></body></html>"""},
                },
            },
        )
        return True
    except Exception as exc:
        st.error(f"Email delivery failed: {exc}")
        return False


# ------------------------------------------------------------------- login --
def _back_to_email(message: str | None = None) -> None:
    st.session_state["login_phase"] = "email"
    st.session_state.pop("pending_otp", None)
    if message:
        st.session_state["login_flash"] = message
    st.rerun()


def _render_login() -> None:
    _, mid, _ = st.columns([1, 2, 1])
    with mid:
        st.markdown(f"## 📊 {APP_NAME}")
        st.caption("Sign in with your SEGA America email to continue.")

        flash = st.session_state.pop("login_flash", None)
        if flash:
            st.error(flash)

        if st.session_state.get("login_phase", "email") == "email":
            email = st.text_input("Email", placeholder="you@segaamerica.com", key="li_email",
                                  label_visibility="collapsed")
            if st.button("Send code", type="primary", width="stretch"):
                email = email.strip().lower()
                if not email.endswith(ALLOWED_DOMAIN):
                    st.error(f"Access is restricted to {ALLOWED_DOMAIN} accounts.")
                else:
                    otp = f"{secrets.randbelow(900000) + 100000}"
                    st.session_state.update({
                        "pending_email": email, "pending_otp": otp,
                        "otp_issued_at": time.time(), "otp_attempts": 0,
                        "login_phase": "otp",
                    })
                    if _send_otp(email, otp):
                        st.rerun()
                    st.session_state["login_phase"] = "email"
        else:
            pending = st.session_state.get("pending_email", "")
            st.info(f"Code sent to **{pending}**")
            code = st.text_input("6-digit code", max_chars=6, key="li_otp",
                                 label_visibility="collapsed", placeholder="000000")
            back, go = st.columns([1, 2])
            if back.button("← Back", width="stretch"):
                _back_to_email()
            if go.button("Verify", type="primary", width="stretch"):
                attempts = st.session_state.get("otp_attempts", 0)
                if attempts >= MAX_ATTEMPTS:
                    _back_to_email("Too many attempts — please request a new code.")
                if time.time() - st.session_state.get("otp_issued_at", 0) > OTP_EXPIRY:
                    _back_to_email("Code expired — please request a new one.")
                if hmac.compare_digest(code.strip(), st.session_state.get("pending_otp", "")):
                    st.session_state["auth_email"] = pending
                    st.query_params["t"] = _make_token(pending)
                    for k in ("pending_otp", "otp_issued_at", "otp_attempts", "login_phase"):
                        st.session_state.pop(k, None)
                    st.rerun()
                st.session_state["otp_attempts"] = attempts + 1
                left = MAX_ATTEMPTS - attempts - 1
                st.error(f"Incorrect code — {left} attempt{'s' if left != 1 else ''} left.")


def _sign_out() -> None:
    st.session_state.pop("auth_email", None)
    st.query_params.pop("t", None)


# ------------------------------------------------------------------ public --
def require_login() -> str:
    """Show the sign-in page and stop the script until the user is signed in.
    Returns the signed-in email."""
    if not _key():
        st.error("COOKIE_SIGNING_KEY is missing from the app's secrets, so sign-in is disabled.")
        st.stop()

    token = st.query_params.get("t", "")
    if token and not st.session_state.get("auth_email"):
        email = _verify_token(token)
        if email:
            st.session_state["auth_email"] = email

    email = st.session_state.get("auth_email")
    if not email:
        _render_login()
        st.stop()

    with st.sidebar:
        st.caption(f"Signed in as **{email}**")
        st.button("Sign out", on_click=_sign_out)
    return email
