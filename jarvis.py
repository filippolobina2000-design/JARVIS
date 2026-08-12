import os
import re
import requests
import win32com.client
import numpy as np
import sounddevice as sd
import speech_recognition as sr
from langdetect import detect, DetectorFactory
from google import genai
from google.genai import types

DetectorFactory.seed = 0  # rende il rilevamento lingua deterministico

NOME_ASSISTENTE = "Jarvis"

# --- LINGUE SUPPORTATE ---
# codice rilevato -> (codice per il riconoscimento vocale, parola chiave voce TTS)
LINGUE = {
    "it": ("it-IT", "italian"),
    "en": ("en-US", "english"),
    "es": ("es-ES", "spanish"),
    "fr": ("fr-FR", "french"),
    "de": ("de-DE", "german"),
    "pt": ("pt-PT", "portuguese"),
}
lingua_riconoscimento = "it-IT"
voce_keyword = "italian"

# --- SINTESI VOCALE (SAPI nativo di Windows: più veloce e affidabile di pyttsx3) ---
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 2  # velocità da -10 (lento) a 10 (veloce); 2 = leggermente più rapido del normale

def imposta_voce(keyword):
    voci = speaker.GetVoices()
    for i in range(voci.Count):
        v = voci.Item(i)
        if keyword in v.GetDescription().lower():
            speaker.Voice = v
            return

def parla(testo):
    imposta_voce(voce_keyword)
    speaker.Speak(testo)

# --- RICONOSCIMENTO VOCALE ---
recognizer = sr.Recognizer()
SAMPLE_RATE = 16000

def registra_audio(max_secondi=10, silenzio_secondi=0.5, soglia=300):
    """Registra dal microfono finché non rileva silenzio dopo che hai parlato."""
    print("🎤 In ascolto...")
    blocco_durata = 0.1
    blocco_campioni = int(SAMPLE_RATE * blocco_durata)
    buffer = []
    iniziato = False
    silenzio_contatore = 0.0
    tempo_totale = 0.0

    with sd.InputStream(samplerate=SAMPLE_RATE, channels=1, dtype="int16") as stream:
        while tempo_totale < max_secondi:
            blocco, _ = stream.read(blocco_campioni)
            buffer.append(blocco.copy())
            volume = np.abs(blocco).mean()

            if volume > soglia:
                iniziato = True
                silenzio_contatore = 0.0
            elif iniziato:
                silenzio_contatore += blocco_durata
                if silenzio_contatore >= silenzio_secondi:
                    break

            tempo_totale += blocco_durata

    if not iniziato:
        return None

    audio_np = np.concatenate(buffer, axis=0)
    audio_bytes = audio_np.tobytes()
    return sr.AudioData(audio_bytes, SAMPLE_RATE, 2)

def ascolta():
    audio = registra_audio()
    if audio is None:
        return None
    try:
        testo = recognizer.recognize_google(audio, language=lingua_riconoscimento)
        print(f"Tu: {testo}")
        return testo
    except sr.UnknownValueError:
        print("(Non ho capito)")
        return ""  # stringa vuota = ho sentito qualcosa ma non l'ho capito
    except sr.RequestError as e:
        print(f"Errore del servizio di riconoscimento vocale: {e}")
        return None

def aggiorna_lingua(testo):
    """Rileva la lingua del testo e aggiorna riconoscimento vocale + voce TTS
    per il turno successivo. Le frasi troppo corte vengono ignorate perché
    il rilevamento automatico non è affidabile su pochi caratteri."""
    global lingua_riconoscimento, voce_keyword
    if len(testo.split()) < 4:
        return
    try:
        codice = detect(testo)
    except Exception:
        return
    if codice in LINGUE:
        lingua_riconoscimento, voce_keyword = LINGUE[codice]

# --- GEOLOCALIZZAZIONE IP ---
# Riconosce sia "8.8.8.8" (scritto) sia "8 8 8 8" (detto a voce, dove il
# riconoscimento vocale spesso non mette i punti)
REGEX_IP = re.compile(r"\b(\d{1,3})[.\s]+(\d{1,3})[.\s]+(\d{1,3})[.\s]+(\d{1,3})\b")

def localizza_ip(indirizzo_ip):
    """Restituisce le informazioni approssimative (città/regione/paese/provider)
    associate a un indirizzo IP, usando il servizio gratuito ip-api.com."""
    try:
        risposta = requests.get(f"http://ip-api.com/json/{indirizzo_ip}?lang=it", timeout=5)
        dati = risposta.json()
    except Exception:
        return "Non sono riuscito a contattare il servizio di geolocalizzazione."

    if dati.get("status") != "success":
        return f"Non sono riuscito a trovare informazioni sull'indirizzo {indirizzo_ip}."

    citta = dati.get("city", "sconosciuta")
    regione = dati.get("regionName", "")
    paese = dati.get("country", "sconosciuto")
    provider = dati.get("isp", "sconosciuto")

    return (
        f"L'indirizzo {indirizzo_ip} risulta registrato a {citta}"
        f"{', ' + regione if regione else ''}, {paese}. "
        f"Il provider associato è {provider}. "
        f"Questa è una stima approssimativa, non una posizione esatta."
    )

# --- METEO (qualsiasi città del mondo, tramite Open-Meteo, gratuito) ---
DESCRIZIONI_METEO = {
    0: "cielo sereno", 1: "prevalentemente sereno", 2: "parzialmente nuvoloso",
    3: "nuvoloso", 45: "nebbia", 48: "nebbia con brina",
    51: "pioviggine leggera", 53: "pioviggine moderata", 55: "pioviggine intensa",
    61: "pioggia leggera", 63: "pioggia moderata", 65: "pioggia intensa",
    71: "nevicata leggera", 73: "nevicata moderata", 75: "nevicata intensa",
    80: "rovesci leggeri", 81: "rovesci moderati", 82: "rovesci violenti",
    95: "temporale", 96: "temporale con grandine", 99: "temporale con grandine intenso",
}

REGEX_METEO = re.compile(
    r"(?:che\s+)?(?:tempo|meteo|previsioni)(?:\s+fa)?(?:\s+(?:a|ad|in|di|per))?\s+([a-zàèéìòù'\- ]+)",
    re.IGNORECASE
)

def get_meteo(citta):
    try:
        geo = requests.get(
            "https://geocoding-api.open-meteo.com/v1/search",
            params={"name": citta, "count": 1, "language": "it", "format": "json"},
            timeout=5
        ).json()
    except Exception:
        return "Non sono riuscito a contattare il servizio meteo."

    risultati = geo.get("results")
    if not risultati:
        return f"Non sono riuscito a trovare la città {citta}."

    luogo = risultati[0]
    lat, lon = luogo["latitude"], luogo["longitude"]
    nome_citta = luogo.get("name", citta)
    paese = luogo.get("country", "")

    try:
        meteo = requests.get(
            "https://api.open-meteo.com/v1/forecast",
            params={"latitude": lat, "longitude": lon, "current_weather": True},
            timeout=5
        ).json()
    except Exception:
        return "Non sono riuscito a ottenere i dati meteo."

    attuale = meteo.get("current_weather")
    if not attuale:
        return f"Non sono riuscito a ottenere il meteo per {nome_citta}."

    temperatura = attuale.get("temperature")
    vento = attuale.get("windspeed")
    codice = attuale.get("weathercode")
    descrizione = DESCRIZIONI_METEO.get(codice, "condizioni non specificate")

    return (
        f"A {nome_citta}, {paese}, al momento ci sono {temperatura}°C, "
        f"{descrizione}, con vento a {vento} km/h."
    )

# --- CONFIGURAZIONE GEMINI ---
# La chiave NON va scritta nel codice. Impostala come variabile d'ambiente.
# Su Windows (Prompt dei comandi): set GEMINI_API_KEY=la-tua-chiave
# Su Windows (PowerShell):         $env:GEMINI_API_KEY="la-tua-chiave"
# Su macOS/Linux (bash/zsh):       export GEMINI_API_KEY="la-tua-chiave"
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")

if not GEMINI_KEY:
    print("ERRORE: variabile d'ambiente GEMINI_API_KEY non trovata.")
    print("Impostala prima di avviare lo script (vedi commento sopra).")
    raise SystemExit(1)

client = genai.Client(api_key=GEMINI_KEY)

# Personalità: si chiama Jarvis, si rivolge con "Signore" (o l'equivalente
# nella lingua in uso), parla come una normale IA, risponde in modo BREVE
# (per essere più rapido da ascoltare) e si adatta alla lingua dell'utente
config = types.GenerateContentConfig(
    system_instruction=(
        "Ti chiami Jarvis. Sei un assistente IA utile, chiaro e diretto. "
        "Rispondi sempre nella stessa lingua in cui ti scrive l'utente, "
        "adattandoti automaticamente se l'utente cambia lingua. Parla in "
        "modo naturale e conciso (massimo 1-2 frasi brevi, a meno che non "
        "ti venga chiesto esplicitamente di approfondire), come farebbe "
        "una normale intelligenza artificiale (non interpretare un "
        "personaggio, non usare un tono teatrale o cinematografico). "
        "Rivolgiti all'utente con un termine rispettoso equivalente a "
        "'Signore' nella lingua in cui state conversando. "
        "IMPORTANTE: non hai accesso a internet, a dati in tempo reale, né "
        "puoi eseguire verifiche, ricerche o localizzazioni per conto tuo. "
        "Non affermare mai di aver 'verificato di nuovo', 'effettuato una "
        "ricerca' o 'localizzato' qualcosa: se ti viene chiesto qualcosa "
        "che richiederebbe dati aggiornati o in tempo reale (meteo, "
        "geolocalizzazione IP, notizie, ecc.) che non hai realmente "
        "ricevuto in questo messaggio, di' chiaramente che non puoi "
        "verificarlo tu direttamente, invece di inventare una risposta "
        "plausibile."
    )
)

print("---------------------------------------")
print(f"        {NOME_ASSISTENTE.upper()} ATTIVO          ")
print("---------------------------------------\n")
messaggio_iniziale = "Ciao Signore, sono pronto. Come posso aiutarla?"
print(f"{NOME_ASSISTENTE}: {messaggio_iniziale}\n")
parla(messaggio_iniziale)

chat = client.chats.create(
    model="gemini-3.5-flash",
    config=config
)

PAROLE_SPEGNI = ["spegni", "disattiva"]
PAROLE_ACCENDI = ["accendi", "attiva", "riattiva"]

microfono_attivo = True

while True:
    try:
        user_input = ascolta()

        if not microfono_attivo:
            # Con il microfono "spento" ascoltiamo comunque, ma ignoriamo
            # tutto tranne il comando per riaccenderlo.
            if user_input:
                t = user_input.lower()
                if "microfono" in t and any(p in t for p in PAROLE_ACCENDI):
                    microfono_attivo = True
                    msg = "Microfono riattivato. Ti ascolto di nuovo."
                    print(f"\n{NOME_ASSISTENTE}: {msg}\n")
                    parla(msg)
            continue

        if user_input is None:
            continue

        if user_input == "":
            # Sentito ma non capito: offriamo la possibilità di scrivere
            testo_scritto = input("Non ho capito. Scrivi qui cosa intendevi (Invio per riprovare a parlare): ")
            if not testo_scritto.strip():
                continue
            user_input = testo_scritto.strip()
            print(f"Tu (scritto): {user_input}")

        testo_minuscolo_check = user_input.lower()
        if "voglio scrivere" in testo_minuscolo_check or "fammi scrivere" in testo_minuscolo_check:
            testo_scritto = input("Scrivi pure qui: ")
            if not testo_scritto.strip():
                continue
            user_input = testo_scritto.strip()
            print(f"Tu (scritto): {user_input}")

        aggiorna_lingua(user_input)
        testo_minuscolo = user_input.lower()

        # --- Comandi speciali ---
        if testo_minuscolo in ["esci", "exit", "quit", "spegniti"]:
            print(f"\n{NOME_ASSISTENTE}: Chiusura in corso. A presto.")
            parla("Chiusura in corso. A presto.")
            break

        if "microfono" in testo_minuscolo and any(p in testo_minuscolo for p in PAROLE_SPEGNI):
            microfono_attivo = False
            msg = "Microfono disattivato. Dì 'accendi il microfono' quando vuoi che ti ascolti di nuovo."
            print(f"\n{NOME_ASSISTENTE}: {msg}\n")
            parla(msg)
            continue

        if "microfono" in testo_minuscolo and any(p in testo_minuscolo for p in PAROLE_ACCENDI):
            msg = "Il microfono è già attivo."
            print(f"\n{NOME_ASSISTENTE}: {msg}\n")
            parla(msg)
            continue

        if "ip" in testo_minuscolo:
            match_ip = REGEX_IP.search(user_input)
            if match_ip:
                indirizzo_ricostruito = ".".join(match_ip.groups())
                msg = localizza_ip(indirizzo_ricostruito)
                print(f"\n{NOME_ASSISTENTE}: {msg}\n")
                parla(msg)
                continue

        match_meteo = REGEX_METEO.search(user_input)
        if match_meteo:
            citta = match_meteo.group(1).strip()
            msg = get_meteo(citta)
            print(f"\n{NOME_ASSISTENTE}: {msg}\n")
            parla(msg)
            continue

        # --- Richiesta normale a Gemini ---
        try:
            response = chat.send_message(user_input)
            testo_risposta = response.text
            if not testo_risposta:
                raise ValueError("Risposta vuota")
        except Exception as e:
            if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e):
                testo_risposta = (
                    "Ho esaurito la quota gratuita dell'API per oggi, Signore. "
                    "Bisogna attendere il reset, di solito il giorno successivo."
                )
            else:
                testo_risposta = (
                    "Non sono riuscito a elaborare una risposta a questa domanda. "
                    "Puoi provare a riformularla?"
                )

        print(f"\n{NOME_ASSISTENTE}: {testo_risposta}\n")
        parla(testo_risposta)

    except KeyboardInterrupt:
        print(f"\n{NOME_ASSISTENTE}: Chiusura manuale. A presto.")
        break
    except Exception as e:
        print(f"\nErrore di sistema: {type(e).__name__}: {e}\n")
        continue
 
