"""Simple user-ID / password protection (hashed passwords, stored in data/users.json)."""
import os, json, secrets, time
from werkzeug.security import generate_password_hash, check_password_hash

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
USERS_FILE, KEY_FILE = os.path.join(DATA, "users.json"), os.path.join(DATA, "secret.key")
os.makedirs(DATA, exist_ok=True)

def load_users():
    try:
        with open(USERS_FILE, encoding="utf-8") as f: return json.load(f)
    except FileNotFoundError:
        return {}

def save_users(u):
    with open(USERS_FILE, "w", encoding="utf-8") as f: json.dump(u, f, indent=2)

def add_user(name, password):
    if len(password) < 8: raise ValueError("Password must be at least 8 characters.")
    u = load_users(); u[name.strip().lower()] = generate_password_hash(password); save_users(u)

def delete_user(name):
    u = load_users(); u.pop(name.strip().lower(), None); save_users(u)

def secret_key():
    if not os.path.exists(KEY_FILE):
        with open(KEY_FILE, "w") as f: f.write(secrets.token_hex(32))
    return open(KEY_FILE).read().strip()

_fails = {}   # ip -> [count, locked_until]
def locked(ip):
    c = _fails.get(ip); return bool(c and c[1] > time.time())
def record(ip, ok):
    if ok: _fails.pop(ip, None); return
    c = _fails.setdefault(ip, [0, 0]); c[0] += 1
    if c[0] >= 5: c[1] = time.time() + 300; c[0] = 0     # 5 wrong tries -> 5 min lock

def verify(name, password):
    h = load_users().get((name or "").strip().lower())
    return bool(h) and check_password_hash(h, password or "")
