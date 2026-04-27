#!/usr/bin/env bash
# new_project.sh — start a new Cortex v2 project
#
# Usage:
#   ./new_project.sh <JOB_NUMBER>  [--from <existing_job>]
#
# Examples:
#   ./new_project.sh 2027                   # blank template
#   ./new_project.sh 2027 --from 2026       # copy 2026's JSON as starting point
#
# What it does:
#   1. Creates job_data/job_data_<JOB>.json (blank template OR copy of another job)
#   2. Prints the 3 commands to finish (fill, build, audit)

set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
JOB="${1:-}"

if [[ -z "$JOB" ]]; then
  echo "Usage: $0 <JOB_NUMBER> [--from <existing_job>]"
  exit 1
fi

TARGET="$HERE/job_data/job_data_${JOB}.json"

if [[ -f "$TARGET" ]]; then
  echo "❌ $TARGET already exists. Delete it first if you really want to start over."
  exit 1
fi

mkdir -p "$HERE/job_data"

if [[ "${2:-}" == "--from" && -n "${3:-}" ]]; then
  SRC="$HERE/job_data/job_data_${3}.json"
  if [[ ! -f "$SRC" ]]; then
    echo "❌ $SRC not found."
    exit 1
  fi
  cp "$SRC" "$TARGET"
  # Update job number inside the copy (works for top-level "job" field)
  python3 -c "
import json, sys
p = '$TARGET'
d = json.loads(open(p).read())
d['job'] = '$JOB'
d['project_name'] = 'REPLACE — copied from $3'
d['source_file'] = '$JOB Job Detail Report.pdf'
open(p, 'w').write(json.dumps(d, indent=2))
"
  echo "✅ Created $TARGET (copied from job $3)"
else
  cp "$HERE/job_data_template.json" "$TARGET"
  python3 -c "
import json, sys
p = '$TARGET'
d = json.loads(open(p).read())
d['job'] = '$JOB'
open(p, 'w').write(json.dumps(d, indent=2))
"
  echo "✅ Created $TARGET from blank template"
fi

echo ""
echo "Next steps:"
echo "  1. Drop source docs into /mnt/owp-${JOB}/ (JDR, contract, billing, etc.)"
echo "  2. Fill in the JSON — either by hand, or ask the agent:"
echo "       \"Read /mnt/owp-${JOB}/ and fill in job_data/job_data_${JOB}.json\""
echo "  3. Build the Excel:"
echo "       python3 $HERE/builder/cortex_builder.py $TARGET \\"
echo "           --out \"$HERE/rebuilt/OWP_${JOB}_JCR_Cortex_v2.xlsx\""
echo "  4. Audit:"
echo "       python3 $HERE/loader/audit.py"
echo ""
