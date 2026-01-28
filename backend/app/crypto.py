from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64
import hashlib
from app.config import get_settings


def _derive_salt(secret_key: str) -> bytes:
    """Derive a deterministic salt from the secret key.

    This ensures each installation has a unique salt based on its SECRET_KEY,
    rather than using a static salt visible in source code.
    """
    return hashlib.sha256(f"ms365-access-salt-derivation:{secret_key}".encode()).digest()


def get_fernet() -> Fernet:
    settings = get_settings()
    secret_key = settings.secret_key
    salt = _derive_salt(secret_key)

    # Derive a key from the secret key using PBKDF2
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=100000,
    )
    key = base64.urlsafe_b64encode(kdf.derive(secret_key.encode()))
    return Fernet(key)


def encrypt_token(token: str) -> str:
    fernet = get_fernet()
    return fernet.encrypt(token.encode()).decode()


def decrypt_token(encrypted_token: str) -> str:
    fernet = get_fernet()
    return fernet.decrypt(encrypted_token.encode()).decode()
