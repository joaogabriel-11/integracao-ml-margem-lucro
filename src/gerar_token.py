import re

from ml_auth import trocar_code_por_token
from ml_auth import REDIRECT_URI
print("REDIRECT_URI em uso:", repr(REDIRECT_URI))

entrada = input("Cole o code TG- (ou a URL inteira): ")

match = re.search(r"TG-[\w-]+", entrada)
if not match:
    raise SystemExit("❌ Não encontrei um código TG- no texto colado.")

code = match.group(0)
print("Código extraído:", code)

tokens = trocar_code_por_token(code)

print("✅ Token gerado!")
print("user_id (seller_id):", tokens["user_id"])
print("Expira em (s):", tokens["expires_in"])
print("Tem refresh_token?:", "refresh_token" in tokens)