import os
from pathlib import Path
from anthropic import Anthropic
from dotenv import load_dotenv
from anonymize import anonymize_text, audit_pii
from rehydrate import rehydrate_text

load_dotenv(Path(__file__).parent / ".env")
client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))

SECURITY_RULES = """
Tu travailles uniquement avec des placeholders anonymisés.

Règles obligatoires :
- Ne jamais modifier les tokens comme {{PERSON_001}}, {{EMAIL_001}}, {{ORG_001}}
- Ne jamais inventer de vraies données personnelles
- Recopier les tokens exactement
- Produire une réponse avec les mêmes tokens
"""

def ask_secure_claude(raw_text):
    safe_text = anonymize_text(raw_text)

    # Garde-fou : bloque si du PII résiduel est détecté après anonymisation
    audit_pii(safe_text)

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        system=SECURITY_RULES,
        messages=[
            {
                "role": "user",
                "content": safe_text
            }
        ]
    )

    safe_answer = message.content[0].text
    final_answer = rehydrate_text(safe_answer)

    return final_answer
