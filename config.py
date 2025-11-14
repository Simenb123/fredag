from datetime import time

# Helgestart (fredag)
WEEKEND_CUTOFF = time(16, 10)

# Norske dagnavn
DAGNAVN = ["mandag", "tirsdag", "onsdag", "torsdag", "fredag", "lørdag", "søndag"]

# Ukesoppsummering
TOP_N_SENDERS = 25

# Ytelses-/sikkerhetsgrenser
MAX_PER_FOLDER = 6000       # maks meldinger vi skanner per mappe manuelt
FALLBACK_RECENT_N = 2000    # fallback: N siste i Innboks (filtreres til uke)

# Hvis Outlook-profilen mangler standardadresse (sjelden) – sett denne manuelt
FALLBACK_EMAIL = None       # f.eks. "din@adresse.no"
