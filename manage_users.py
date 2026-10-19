"""
  python manage_users.py add alice        (asks for the password)
  python manage_users.py delete alice
  python manage_users.py list
"""
import sys, getpass, auth
cmd = sys.argv[1] if len(sys.argv) > 1 else ""
if cmd == "add" and len(sys.argv) > 2:
    p1 = getpass.getpass("Password (min 8 chars): "); p2 = getpass.getpass("Repeat: ")
    if p1 != p2: sys.exit("Passwords do not match.")
    auth.add_user(sys.argv[2], p1); print("User saved.")
elif cmd == "delete" and len(sys.argv) > 2:
    auth.delete_user(sys.argv[2]); print("Deleted.")
elif cmd == "list":
    print("\n".join(auth.load_users()) or "(no users)")
else:
    print(__doc__)
