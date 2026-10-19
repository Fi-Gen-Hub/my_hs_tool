HS CODE TOOL - quick start (PyCharm)
1. pip install -r requirements.txt
2. python build_database.py          (one time, needs internet, resumable, builds the OFFLINE database)
3. python build_database.py --verify (optional: spot-checks 25 random codes DB vs live website)
4. python app.py  -> open http://localhost:5000
   Data Source: Auto = offline DB first, live site if missing | Offline only | Live only

LOGIN
  python manage_users.py add yourname     (create user ID + password, min 8 chars)
  python manage_users.py list | delete yourname
  The app refuses to start until at least one user exists.
  Default is localhost only. To allow other PCs on your network: set HS_HOST=0.0.0.0
  (use HTTPS via a reverse proxy if exposed beyond your office, and set HS_HTTPS=1).

PRIVATE GIT REPOSITORY
  .gitignore already excludes data/ (users, password hashes, secret key, database) and uploads/.
  git init && git add . && git commit -m "HS code tool"
  Create a PRIVATE repo on GitHub (or GitLab/Bitbucket), then:
  git remote add origin <repo-url> && git push -u origin main
