set -e
S=${MA_WORK}
K=/home/user/github-1st/.claude/skills/ma-deal-docs/scripts
cd $S && python3 build_master.py
cd $S/master && rm -rf recalc && timeout 180 soffice --headless --norestore -env:UserInstallation=file:///tmp/lo-calc --convert-to xlsx --outdir recalc master_raw.xlsx >/dev/null 2>&1
python3 $K/vr_prep.py master_raw.xlsx recalc/master_raw.xlsx master_raw2.xlsx
python3 $K/inject_cache.py master_raw2.xlsx recalc/master_raw.xlsx master.xlsx
