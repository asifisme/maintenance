import os
import django
from django.core.management import call_command

os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")
django.setup()

print("--- Running makemigrations ---")
try:
    call_command('makemigrations', interactive=False)
except Exception as e:
    print("Makemigrations error:", e)

print("--- Running migrate ---")
try:
    call_command('migrate', interactive=False)
    print("Migrate complete successfully.")
except Exception as e:
    print("Migrate error:", e)
