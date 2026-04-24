"""Microsoft Graph client using access token from Graph Explorer.

Token is resolved in order: persisted file → GRAPH_ACCESS_TOKEN env var.
"""

import asyncio
import os
from pathlib import Path
from azure.core.credentials import AccessToken, TokenCredential
from msgraph import GraphServiceClient

TOKEN_ENV_VAR = "GRAPH_ACCESS_TOKEN"
TOKEN_FILE = Path.home() / ".config" / "claude-graph" / "token"

TOKEN_INSTRUCTIONS = """
To get an access token:

1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Click the person icon in the top right to sign in
3. Click the "Access token" tab in the request section
4. Click the copy button to copy the token
5. Set the environment variable:

   export GRAPH_ACCESS_TOKEN='<paste token here>'

Note: Tokens expire after ~1 hour. Get a fresh one when needed.
"""


class StaticTokenCredential(TokenCredential):
    """Credential that uses a static access token."""

    def __init__(self, token: str):
        self._token = token

    def get_token(self, *scopes, **kwargs) -> AccessToken:
        return AccessToken(self._token, 0)


def get_token() -> str | None:
    """Get token from persisted file or environment variable.

    File is checked first because it always contains the most recently
    persisted token (written by set_token after auth). The env var may
    be stale if auth ran in a subprocess.
    """
    try:
        if TOKEN_FILE.exists():
            token = TOKEN_FILE.read_text().strip()
            if token:
                return token
    except OSError:
        pass
    token = os.environ.get(TOKEN_ENV_VAR)
    if token:
        return token
    return None


def print_token_instructions() -> None:
    """Print instructions for getting a token."""
    print(TOKEN_INSTRUCTIONS)


def set_token(token: str) -> None:
    """Set token in environment variable and persist to disk."""
    token = token.strip()
    os.environ[TOKEN_ENV_VAR] = token
    try:
        TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(token + "\n")
        TOKEN_FILE.chmod(0o600)
    except OSError:
        pass


def is_token_expired(token: str = None) -> bool:
    """Check if a JWT token is expired by decoding the payload.

    Uses a 60-second buffer to catch tokens that are about to expire.
    Returns False if the token can't be decoded (let the API call determine validity).
    """
    import base64
    import json
    from datetime import datetime, timezone

    if token is None:
        token = get_token()
    if not token:
        return True

    try:
        parts = token.split(".")
        if len(parts) != 3:
            return False

        # Decode base64url payload
        payload = parts[1]
        payload += "=" * (4 - len(payload) % 4)
        decoded = base64.urlsafe_b64decode(payload)
        claims = json.loads(decoded)

        exp = claims.get("exp")
        if exp is None:
            return False

        return datetime.now(timezone.utc).timestamp() > (exp - 60)
    except Exception:
        return False


def get_graph_client(token: str = None) -> GraphServiceClient:
    """Get GraphServiceClient using an access token.

    Args:
        token: Access token. If not provided, resolved via get_token().

    Returns:
        Authenticated GraphServiceClient ready for API calls.
    """
    if token is None:
        token = get_token()
        if token is None:
            raise ValueError(f"No token found. {TOKEN_INSTRUCTIONS}")

    credential = StaticTokenCredential(token)
    return GraphServiceClient(credentials=credential)


def get_concurrency_semaphore(limit: int = 5) -> asyncio.Semaphore:
    """Return an asyncio.Semaphore to cap concurrent Graph API calls.

    Use within a single asyncio.run() scope to limit burst size when
    firing multiple requests via asyncio.gather().

    Args:
        limit: Maximum concurrent in-flight requests (default: 5).
    """
    return asyncio.Semaphore(limit)


def run_async(coro):
    """Run an async coroutine synchronously."""
    return asyncio.run(coro)


def verify_token(client: GraphServiceClient = None) -> bool:
    """Test if the current token is valid by making a simple API call.

    Args:
        client: Optional GraphServiceClient. If not provided, creates one.

    Returns:
        True if token is valid, False otherwise.
    """
    from msgraph.generated.models.o_data_errors.o_data_error import ODataError

    try:
        if client is None:
            client = get_graph_client()
        me = run_async(client.me.get())
        print(f"Token valid. Signed in as: {me.display_name} ({me.user_principal_name})")
        return True
    except ODataError as e:
        print(f"Token invalid or expired: {e.error.message}")
        print_token_instructions()
        return False
    except ValueError as e:
        print(str(e))
        return False
