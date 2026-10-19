"""
Command line helper.   Run from the my_hs_tool folder:

  python build_database.py                 # download the whole PDK tariff (resumable)
  python build_database.py --tariff RCEP   # another tariff
  python build_database.py --verify        # spot-check DB against the live site
  python build_database.py --stats         # what is in the DB
"""
import argparse, hs_core

ap = argparse.ArgumentParser()
ap.add_argument("--tariff", default="PDK", choices=list(hs_core.TARIFF_TYPES))
ap.add_argument("--workers", type=int, default=4)
ap.add_argument("--verify", action="store_true")
ap.add_argument("--stats", action="store_true")
a = ap.parse_args()

if a.stats:
    for s in hs_core.db_stats(): print(s)
elif a.verify:
    hs_core.verify_database(a.tariff)
else:
    hs_core.build_database(a.tariff, a.workers,
        progress=lambda d, t: print("\r  %d / %d" % (d, t), end="", flush=True))
    print()
    for s in hs_core.db_stats(): print(s)
