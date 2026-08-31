import os
import secrets
from fastapi import Header, HTTPException, Depends
from typing import Optional

# API key is read from environment variable, or auto-generated on first run
# For production, set ADMIN_API_KEY in your environment
API_KEY = os.environ.get("ADMIN_API_KEY", "")

_key_generated = False


def get_api_key() -> str:
    """Get the configured API key."""
    global _key_generated
    global API_KEY
    if not API_KEY and not _key_generated:
        API_KEY = secrets.token_urlsafe(32)
        _key_generated = True
        print(f"\n{'='*60}")
        print(f"  ADMIN API KEY (auto-generated): {API_KEY}")
        print(f"  Set ADMIN_API_KEY env var for persistent key")
        print(f"{'='*60}\n")
    return API_KEY


async def verify_admin_key(
    x_admin_key: Optional[str] = Header(None)
) -> Optional[str]:
    """FastAPI dependency: verify the admin API key from header.

    Pass X-Admin-Key header with your API key to authenticate write requests.
    If ADMIN_API_KEY env is not set, access is open (dev mode).
    """
    expected = get_api_key()
    if not expected:
        return ""
    if not x_admin_key or not secrets.compare_digest(x_admin_key, expected):
        raise HTTPException(
            status_code=401,
            detail="Unauthorized: Invalid or missing X-Admin-Key header"
        )
    return x_admin_key
