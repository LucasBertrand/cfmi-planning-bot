import os
import discord
import aiohttp
from discord.ext import commands, tasks
from discord import app_commands
from datetime import date, time, datetime, timedelta
from typing import Optional, Dict, Any, Union
from collections import defaultdict
import openpyxl
import tempfile
import json
from dataclasses import dataclass
import uuid
from dotenv import load_dotenv
import unicodedata

load_dotenv()

# --- CONSTANTS ---

BOT_TOKEN = os.getenv("BOT_TOKEN")
CHANNEL_ID = int(os.getenv("CHANNEL_ID"))

BIRTHDAY_URL = "birthdays.json"
PLANNING_URL = "https://sha.univ-poitiers.fr/cfmi/wp-content/uploads/sites/276/2025/10/Planning-general-CFMI-25-26.xlsx"
DEFAULT_PROMO = "2A"
EXCLUDED_MODULES = ["CONGES UNIVERSITAIRES", "FERIE"]

FRENCH_MONTHS = {
  "JANVIER": 1, 
  "FEVRIER": 2, 
  "MARS": 3, 
  "AVRIL": 4,
  "MAI": 5, 
  "JUIN": 6, 
  "JUILLET": 7, 
  "AOUT": 8,
  "SEPTEMBRE": 9, 
  "OCTOBRE": 10, 
  "NOVEMBRE": 11, 
  "DECEMBRE": 12,
}

COLS = {
  "DAY": 0, 
  "WEEKDAY": 1, 
  "MONTH": 2, 
  "YEAR": 3, 
  "PROMO": 4,
  "TYPE": 5, 
  "START_TIME": 6, 
  "END_TIME": 7, 
  "MODULE_CODE": 8,
  "MODULE_NAME": 9, 
  "PROFESSOR": 10, 
  "LOCATION": 11, 
  "GROUP": 12,
  "SESSION": 13, 
  "NOTE": 14,
}

# --- ERROR MESSAGES ---

ERROR_INVALID_DATE = "Oups ! Date invalide. Utilise l'un des formats : JJ, JJ/MM, JJ/MM/AA ou JJ/MM/AAAA."
ERROR_TIME_ORDER = "Oups ! L'heure de fin doit être après l'heure de début."
ERROR_INVALID_TIME = "Oups ! Heure invalide. Utilise le format HH ou HH:MM."

# --- BIRTHDAY MESSAGES ---

BIRTHDAY_LIST_TITLE = "🎂  Liste des anniversaires"
BIRTHDAY_LIST_DESCRIPTION = "Affiche tous les anniversaires enregistrés."
BIRTHDAY_ADD_DESCRIPTION = "Ajoute ton anniversaire pour qu'on puisse le fêter !"
BIRTHDAY_ADD_OK = "Ton anniversaire a été enregistré !"
BIRTHDAY_DELETE_OK = "Ton anniversaire a été supprimé."
BIRTHDAY_DELETE_DESCRIPTION = "Supprime ton anniversaire du calendrier."
BIRTHDAY_NONE = "Aucun anniversaire enregistré pour le moment."
BIRTHDAY_NOTIFICATION = "vient d'ajouter son anniversaire !"

# --- PLANNING MESSAGES ---

TODAY_DESCRIPTION = "Affiche les événements d'aujourd'hui."
TOMORROW_DESCRIPTION = "Affiche les événements de demain."
NEXT_DESCRIPTION = "Affiche les événements du prochain jour de la semaine indiqué."
PLANNING_DESCRIPTION = "Affiche les événements prévus pour une date donnée."
PLANNING_TITLE = "Au programme du "
PLANNING_NONE = "Aucun événement prévu pour cette date."

# --- SEARCH MESSAGES ---

SEARCH_DESCRIPTION = "Recherche un événement dans le planning."
SEARCH_TOO_SHORT = "La recherche doit contenir au moins 3 caractères."
SEARCH_NO_RESULTS = "Aucun événement n'a été trouvé."
SEARCH_RESULTS_TITLE = "🔎  Résultats de recherche"
SEARCH_RESULTS_LIMIT = "⚠️  Résultats limités à 15. Affine peut-être ta recherche."

# --- PARAMETER DESCRIPTIONS ---

DATE_DESCRIPTION = "La date de l’événement (formats : JJ, JJ/MM, JJ/MM/AA ou JJ/MM/AAAA)"
PROMO_DESCRIPTION = "Choisis la promotion (1A, 2A ou les deux)"
JOUR_DESCRIPTION = "Choisis le jour de la semaine pour lequel afficher le planning"
REQUETE_DESCRIPTION = "Liste de mots clés (au moins 3 caractères, ignore les accents)"
DEBUT_DESCRIPTION = "Date de début de la recherche (optionnelle, par défaut aujourd’hui)"
FIN_DESCRIPTION = "Date de fin de la recherche (optionnelle, par défaut dernière date du planning)"

# --- DATA CLASSES ---

@dataclass
class PlanningEvent:
  event_date: date
  organizer: Optional[Union[str, int]] = None
  title: Optional[str] = None
  start_time: Optional[time] = None
  end_time: Optional[time] = None
  location: Optional[str] = None
  note: Optional[str] = None
  promo: Optional[int] = None
  module: Optional[str] = None
  def to_dict(self) -> Dict[str, Any]:
    return {
      "event_date": self.event_date.isoformat(),
      "organizer": self.organizer,
      "title": self.title,
      "start_time": self.start_time.isoformat() if self.start_time else None,
      "end_time": self.end_time.isoformat() if self.end_time else None,
      "location": self.location,
      "note": self.note,
      "promo": self.promo,
      "module": self.module
    }
  @classmethod
  def from_dict(cls, data: Dict[str, Any]) -> "PlanningEvent":
    """Construct from dict produced by to_dict (or raw loaded JSON)."""
    ed = data.get("event_date")
    if isinstance(ed, str): ed = date.fromisoformat(ed)
    st = data.get("start_time")
    if isinstance(st, str): st = time.fromisoformat(st)
    et = data.get("end_time")
    if isinstance(et, str): et = time.fromisoformat(et)
    return cls(
      event_date = ed,
      organizer = data.get("organizer"),
      title = data.get("title"),
      start_time = st,
      end_time = et,
      location = data.get("location"),
      note = data.get("note"),
      promo = data.get("promo"),
      module = data.get("module")
    )
  
# --- UTILITIES ---

def normalize_str(s: str) -> str:
  """Lowercase and remove accents from a string."""
  if not s: return ""
  nfd_form = unicodedata.normalize("NFD", s)
  return "".join([c for c in nfd_form if unicodedata.category(c) != "Mn"]).lower()

def generate_uid() -> str:
  """Generate and return a unique identifier as a string."""
  return str(uuid.uuid4())

def log(message: str):
  """Log a message with a timestamp."""
  timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  print(f"{timestamp} - {message}")

async def load_json_file(path: str) -> dict:
  """Load and return the JSON content of a file."""
  with open(path, "r") as f: return json.load(f)

async def save_json_file(path: str, data: dict):
  """Save data as JSON to a file."""
  with open(path, "w") as f: json.dump(data, f, indent=2)

def parse_time(value: str) -> Optional[time]:
  """Parse a time string in HH or HH:MM format and return a time object."""
  for fmt in ("%H:%M", "%H"):
    try: return datetime.strptime(value, fmt).time()
    except ValueError: continue
  return None

def toStrTime(t: time) -> str:
  """Convert a time object to a string"""
  return f"{t.hour}:{t.minute:02d}"

def parse_date(value: str) -> Optional[date]:
  """Parse a date string in various formats and return a date object."""
  today = date.today()
  patterns = [
    ("%d/%m/%Y", {}),
    ("%d/%m/%y", {}),
    ("%d/%m", {"year": today.year}),
    ("%d", {"month": today.month, "year": today.year}),
  ]
  for fmt, defaults in patterns:
    try:
      dt = datetime.strptime(value, fmt)
      return date(
        year=defaults.get("year", dt.year),
        month=defaults.get("month", dt.month),
        day=dt.day,
      )
    except ValueError: continue
  return None

def get_delta_days(target_weekday: int) -> int:
  """Calculate days until the next occurrence of the target weekday."""
  today = date.today()
  today_weekday = today.weekday()
  delta_days = (target_weekday - today_weekday) % 7
  return delta_days if delta_days != 0 else 7

# --- BIRTHDAYS EVENTS ---

async def load_birthdays():
  """Load birthday events from the JSON file into birthdays cache."""  
  global birthdays_cache
  data = await load_json_file(BIRTHDAY_URL)
  birthdays_cache = {int(k): date.fromisoformat(v) for k,v in data.items()}
  log("Birthdays loaded.")

async def save_birthday(uid: int, birthday: date):
  """Save or update a user's birthday in JSON file and cache."""
  birthdays_cache[uid] = birthday
  await save_json_file(BIRTHDAY_URL, {k: v.isoformat() for k,v in birthdays_cache.items()})

async def remove_birthday(uid: int):
  """Remove a user's birthday from JSON file and cache."""
  if uid in birthdays_cache:
    del birthdays_cache[uid]
    await save_json_file(BIRTHDAY_URL, {k: v.isoformat() for k,v in birthdays_cache.items()})

def get_birthdays(date: date) -> list[tuple[int, date]]:
  """Return a list of (user_id, birthday) tuples for users whose birthday is on the given date."""
  return [(uid, bday) for uid, bday in birthdays_cache.items() if bday.day == date.day and bday.month == date.month]

async def get_annonce_birthday_field(uid: int, birthday: date) -> Dict[str, Any]:
  """Create and return an embed field for a birthday."""
  user = await bot.fetch_user(uid)
  return {"name": f"🎂  Anniversaire de {user.display_name} !", "value": "", "inline": False}

async def get_list_birthday_field(uid: int, birthday: date) -> Dict[str, Any]:
  """Create and return an embed field for a birthday."""
  user = await bot.fetch_user(uid)  
  return {"name": "", "value": f"{user.display_name} - {birthday.strftime('%d/%m/%Y')}", "inline": False}

# --- PLANNING EVENTS ---

async def download_excel(url: str) -> str:
  """Download an Excel file asynchronously from a URL and return the local file path."""
  tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
  async with aiohttp.ClientSession() as session:
    async with session.get(url) as response:
      response.raise_for_status()
      while True:
        chunk = await response.content.read(8192)
        if not chunk: break
        tmp_file.write(chunk)
  tmp_file.close()
  return tmp_file.name

def extract_date_from_row(row) -> Optional[date]:
  """Extract a date object from a row if valid, else return None."""
  day, month_name, year = row[0], row[2], row[3]
  month = FRENCH_MONTHS.get(month_name.upper())
  if all(isinstance(x, int) for x in [day, month, year]):
    return date(year, month, day)
  return None

def has_planning_event(row) -> bool:
  """Check if a row contains a valid planning event."""
  return (
    isinstance(row[COLS.get("MODULE_NAME")], str) and
    row[COLS.get("MODULE_NAME")] not in EXCLUDED_MODULES
  )

def cache_planning_event(row, event_date) -> None:
  """Add a planning event from a row to the planning cache."""
  event = PlanningEvent(
    title = row[COLS['MODULE_NAME']],
    event_date = event_date,
    start_time = row[COLS["START_TIME"]],
    end_time = row[COLS["END_TIME"]],
    location = row[COLS["LOCATION"]],
    organizer = row[COLS["PROFESSOR"]],
    note = row[COLS["NOTE"]],
    promo = row[COLS["PROMO"]],
    module = row[COLS["MODULE_CODE"]]
  )
  planning_cache[event_date].append(event)

@tasks.loop(time=time(2, 0, 0))
async def load_planning_events() -> None:
  """Load planning events from the Excel file into the planning cache."""  
  local_file = await download_excel(PLANNING_URL)
  workbook = openpyxl.load_workbook(local_file, data_only=True, read_only=True)  
  planning_cache.clear()
  for row in workbook.active.iter_rows(min_row=2, values_only=True):
    event_date = extract_date_from_row(row)
    if event_date and has_planning_event(row):
      cache_planning_event(row, event_date)
  workbook.close()
  os.remove(local_file)
  log("Planning events loaded.")

def filter_by_promo(events: list[PlanningEvent], promo: str) -> list[PlanningEvent] :
  """Filter events by promotion."""
  if not events: return []
  if not promo: return events
  return [ev for ev in events if is_in_promo(ev, promo)]

def is_in_promo(event: PlanningEvent, promo: str) -> bool:
  """Check if an event matches the given promotion."""
  if not promo: return True
  p = event.promo
  return p is None or p == promo or (p == "1A/2A" and promo in ["1A", "2A"])

def get_planning_event_field(event: PlanningEvent, with_date: bool = False) -> dict[str, Any]:
  """Create and return an embed field for a planning event."""
  content = ""
  if with_date:
    content += f"- Date : {event.event_date.strftime('%d/%m/%Y')}"
  if event.start_time:
    content += f"\n- Horaire : {toStrTime(event.start_time)}"
  if event.end_time:
    content += f" - {toStrTime(event.end_time)}"
  if event.location:
    content += f"\n- Lieu : {event.location}"
  if event.organizer:
    content += f"\n- Professeur : {event.organizer}"  
  if event.note:
    content += f"\n- Note : {event.note}"
  return {
    "name" : f"📚  {f"{event.promo} - " if event.promo else ""} {event.title}",
    "value" : content,
    "inline" : False
  }

@tasks.loop(time=time(18, 0, 0))
async def daily_tomorrow_planning():
  """Send tomorrow's planning and birthdays to the planning channel."""
  event_date = date.today() + timedelta(days=1)
  channel = bot.get_channel(CHANNEL_ID)
  planning_events = planning_cache.get(event_date)
  birthdays = get_birthdays(event_date)
  if (not planning_events) and (not birthdays): return
  embed = discord.Embed(title=PLANNING_TITLE + event_date.strftime('%d/%m/%Y'), color=discord.Color.blue())  
  for event in birthdays: embed.add_field(**(await get_annonce_birthday_field(*event)))
  for event in planning_events: embed.add_field(**get_planning_event_field(event))
  await channel.send(embed=embed)

# --- GLOBAL CACHES ---

planning_cache: Dict[date, list[PlanningEvent]] = defaultdict(list)
birthdays_cache: Dict[int, date] = {}

# --- DISCORD BOT SETUP ---

intents = discord.Intents.default()
intents.message_content = True
intents.messages = True
intents.members = True
intents.dm_messages = True
bot = commands.Bot(command_prefix="!", intents=intents)

@bot.event
async def on_ready():
  """Actions to perform when the bot is ready."""
  log(f"Bot connected as {bot.user}")  
  load_planning_events.start()
  daily_tomorrow_planning.start()
  await load_birthdays()  
  await load_planning_events()
  await bot.tree.sync()

# --- BIRTHDAY COMMANDS ---

@bot.tree.command(name="list_birthdays", description=BIRTHDAY_LIST_DESCRIPTION)
async def list_birthdays_cmd(interaction: discord.Interaction):
  """List all registered birthdays."""
  birthdays = list(birthdays_cache.items())
  if not birthdays: return await interaction.response.send_message(BIRTHDAY_NONE, ephemeral=True)
  embed = discord.Embed(title=BIRTHDAY_LIST_TITLE, color=discord.Color.blue())
  for uid, birthday in birthdays: embed.add_field(**(await get_list_birthday_field(uid, birthday)))
  await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="add_birthday", description=BIRTHDAY_ADD_DESCRIPTION)
async def add_birthday_cmd(interaction: discord.Interaction, date_de_naissance: str):
  """Add or update user's birthday."""
  channel = bot.get_channel(CHANNEL_ID)
  birthday = parse_date(date_de_naissance)
  if not birthday or (date.today() < birthday): 
    return await interaction.response.send_message(ERROR_INVALID_DATE, ephemeral=True)
  await save_birthday(interaction.user.id, birthday)
  await interaction.response.send_message(BIRTHDAY_ADD_OK, ephemeral=True)
  await channel.send(f"🎉  {interaction.user.display_name} {BIRTHDAY_NOTIFICATION} ({birthday.strftime('%d/%m/%Y')})")

@bot.tree.command(name="del_birthday", description=BIRTHDAY_DELETE_DESCRIPTION)
async def delete_birthday_cmd(interaction: discord.Interaction):
  """Delete user's birthday."""
  await remove_birthday(interaction.user.id)
  await interaction.response.send_message(BIRTHDAY_DELETE_OK, ephemeral=True)

# --- PLANNING EVENT COMMAND ---

@bot.tree.command(name="today", description=TODAY_DESCRIPTION)
@app_commands.describe(promotion=PROMO_DESCRIPTION)
@app_commands.choices(promotion=[
  app_commands.Choice(name="1A/2A", value=""), 
  app_commands.Choice(name="1A", value="1A"), 
  app_commands.Choice(name="2A", value="2A")
])
async def today_cmd(interaction: discord.Interaction, promotion: str = ""):
  """List planning events for today."""
  event_date = date.today()
  planning_events = filter_by_promo(planning_cache.get(event_date), promotion)
  birthdays = get_birthdays(event_date)
  if not (planning_events + birthdays): return await interaction.response.send_message(PLANNING_NONE, ephemeral=True)
  embed = discord.Embed(title=PLANNING_TITLE + event_date.strftime('%d/%m/%Y'), color=discord.Color.blue())  
  for event in birthdays: embed.add_field(**(await get_annonce_birthday_field(*event)))
  for event in planning_events: embed.add_field(**get_planning_event_field(event))
  await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="tomorrow", description=TOMORROW_DESCRIPTION)
@app_commands.describe(promotion=PROMO_DESCRIPTION)
@app_commands.choices(promotion=[
  app_commands.Choice(name="1A/2A", value=""), 
  app_commands.Choice(name="1A", value="1A"), 
  app_commands.Choice(name="2A", value="2A")
])
async def tomorrow_cmd(interaction: discord.Interaction, promotion: str = ""):
  """List planning events for tomorrow."""
  event_date = date.today() + timedelta(days=1)
  planning_events = filter_by_promo(planning_cache.get(event_date), promotion)
  birthdays = get_birthdays(event_date)
  if not (planning_events + birthdays): return await interaction.response.send_message(PLANNING_NONE, ephemeral=True)
  embed = discord.Embed(title=PLANNING_TITLE + event_date.strftime('%d/%m/%Y'), color=discord.Color.blue())  
  for event in birthdays: embed.add_field(**(await get_annonce_birthday_field(*event)))
  for event in planning_events: embed.add_field(**get_planning_event_field(event))
  await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="planning", description=PLANNING_DESCRIPTION)
@app_commands.describe(date=DATE_DESCRIPTION, promotion=PROMO_DESCRIPTION)
@app_commands.choices(promotion=[
  app_commands.Choice(name="1A/2A", value=""), 
  app_commands.Choice(name="1A", value="1A"), 
  app_commands.Choice(name="2A", value="2A")
])
async def planning_cmd(interaction: discord.Interaction, date: str, promotion: str = ""):
  """List planning events for a given date."""
  event_date = parse_date(date)
  if not event_date: return await interaction.response.send_message(ERROR_INVALID_DATE, ephemeral=True)
  planning_events = filter_by_promo(planning_cache.get(event_date), promotion)
  birthdays = get_birthdays(event_date)
  if not (planning_events + birthdays): return await interaction.response.send_message(PLANNING_NONE, ephemeral=True)
  embed = discord.Embed(title=PLANNING_TITLE + event_date.strftime('%d/%m/%Y'), color=discord.Color.blue())  
  for event in birthdays: embed.add_field(**(await get_annonce_birthday_field(*event)))
  for event in planning_events: embed.add_field(**get_planning_event_field(event))
  await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="next", description=NEXT_DESCRIPTION)
@app_commands.describe(
  jour = JOUR_DESCRIPTION,
  promotion = PROMO_DESCRIPTION
)
@app_commands.choices(promotion=[
  app_commands.Choice(name="1A/2A", value=""), 
  app_commands.Choice(name="1A", value="1A"), 
  app_commands.Choice(name="2A", value="2A")
])
@app_commands.choices(jour=[
  app_commands.Choice(name="lundi", value=0),
  app_commands.Choice(name="mardi", value=1),
  app_commands.Choice(name="mercredi", value=2),
  app_commands.Choice(name="jeudi", value=3),
  app_commands.Choice(name="vendredi", value=4),
  app_commands.Choice(name="samedi", value=5),
  app_commands.Choice(name="dimanche", value=6), 
])
async def next_cmd(interaction: discord.Interaction, jour: int, promotion: str = ""):
  """List planning events for the next given weekday."""
  event_date = date.today() + timedelta(days=get_delta_days(jour))
  planning_events = filter_by_promo(planning_cache.get(event_date), promotion)
  birthdays = get_birthdays(event_date)
  if not (planning_events + birthdays):
    return await interaction.response.send_message(PLANNING_NONE, ephemeral=True)
  embed = discord.Embed(title=PLANNING_TITLE + event_date.strftime('%d/%m/%Y'), color=discord.Color.blue())  
  for event in birthdays: embed.add_field(**(await get_annonce_birthday_field(*event)))
  for event in planning_events: embed.add_field(**get_planning_event_field(event))
  await interaction.response.send_message(embed=embed, ephemeral=True)

@bot.tree.command(name="search", description=SEARCH_DESCRIPTION)
@app_commands.describe(
  requête = REQUETE_DESCRIPTION,
  promotion = PROMO_DESCRIPTION,
  début = DEBUT_DESCRIPTION,
  fin = FIN_DESCRIPTION
)
@app_commands.choices(promotion=[
  app_commands.Choice(name="1A/2A", value=""), 
  app_commands.Choice(name="1A", value="1A"), 
  app_commands.Choice(name="2A", value="2A")
])
async def search_planning_cmd(
  interaction: discord.Interaction,
  requête: str,
  promotion: str = "",
  début: str = None,
  fin: str = None
):
  """Search planning events between optional dates with a query."""
  limit = 15
  q = requête.strip().lower()
  if len(q) < 3:
    return await interaction.response.send_message(SEARCH_TOO_SHORT, ephemeral=True)
  keywords = [normalize_str(word) for word in q.split()]
  start = parse_date(début) if début else date.today()
  end = parse_date(fin) if fin else None
  if début and not start:
    return await interaction.response.send_message(ERROR_INVALID_DATE, ephemeral=True)
  if fin and not end:
    return await interaction.response.send_message(ERROR_INVALID_DATE, ephemeral=True)
  if not end and planning_cache:
    end = max(planning_cache.keys())
  results: list[PlanningEvent] = []
  for d in sorted(planning_cache.keys()):
    if d < start or d > end: continue
    for ev in planning_cache[d]:
      if not is_in_promo(ev, promotion): continue
      haystacks = " ".join([
        str(ev.title or ""),
        str(ev.organizer or ""),
        str(ev.note or "")
      ])
      haystacks = normalize_str(haystacks)
      if all(word in haystacks for word in keywords):
        results.append(ev)
  if not results: return await interaction.response.send_message(SEARCH_NO_RESULTS, ephemeral=True)
  limited = results[:limit]            
  embed = discord.Embed(title=SEARCH_RESULTS_TITLE, color=discord.Color.blue())
  for ev in limited: embed.add_field(**get_planning_event_field(ev, with_date=True))
  if len(results) > limit: embed.set_footer(text=SEARCH_RESULTS_LIMIT)
  await interaction.response.send_message(embed=embed, ephemeral=True)

# --- RUN BOT ---

bot.run(BOT_TOKEN)