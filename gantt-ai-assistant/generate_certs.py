"""
Generate a self-signed SSL certificate for localhost development.
Outputs:
  certs/localhost.key
  certs/localhost.crt
"""
import ipaddress
import os
from datetime import datetime, timezone, timedelta
from pathlib import Path
from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import NameOID

CERT_DIR = Path(__file__).parent / "certs"
CERT_DIR.mkdir(exist_ok=True)

KEY_FILE  = CERT_DIR / "localhost.key"
CERT_FILE = CERT_DIR / "localhost.crt"

# Generate key
private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)

# Write private key
KEY_FILE.write_bytes(
    private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    )
)

# Build certificate
name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
now  = datetime.now(timezone.utc)

cert = (
    x509.CertificateBuilder()
    .subject_name(name)
    .issuer_name(name)
    .public_key(private_key.public_key())
    .serial_number(x509.random_serial_number())
    .not_valid_before(now)
    .not_valid_after(now + timedelta(days=825))
    .add_extension(
        x509.SubjectAlternativeName([
            x509.DNSName("localhost"),
            x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
        ]),
        critical=False,
    )
    .add_extension(
        x509.BasicConstraints(ca=True, path_length=None),
        critical=True,
    )
    .sign(private_key, hashes.SHA256())
)

CERT_FILE.write_bytes(cert.public_bytes(serialization.Encoding.PEM))

print(f"✅ Certificates generated:")
print(f"   Key : {KEY_FILE}")
print(f"   Cert: {CERT_FILE}")
print()
print("Next step: trust this certificate in Windows to avoid browser warnings.")
print("Run this in an admin PowerShell:")
print(f'  Import-Certificate -FilePath "{CERT_FILE}" -CertStoreLocation Cert:\\LocalMachine\\Root')
