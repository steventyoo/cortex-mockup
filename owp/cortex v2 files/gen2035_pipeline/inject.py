#!/usr/bin/env python3
"""Inject generated JS blocks into index.html at the correct anchor points."""
from pathlib import Path

INDEX = Path('/sessions/gracious-relaxed-pascal/mnt/cortex-mockup/index.html')
src = INDEX.read_text()

projects_block = Path('/sessions/gracious-relaxed-pascal/gen2035/projects_block.js').read_text()
teams_block = Path('/sessions/gracious-relaxed-pascal/gen2035/teams_block.js').read_text()
arrays_block = Path('/sessions/gracious-relaxed-pascal/gen2035/arrays_block.js').read_text()

# 1) Inject PROJECTS entries after the 2034 entry closing.
# Anchor: The 2034 object closes at `    }\n  };` (end of PROJECTS assignment)
# Find the unique '  };\n\n  /* ============ PROJECT TEAMS'
anchor_p = '    }\n  };\n\n  /* ============ PROJECT TEAMS'
assert src.count(anchor_p) == 1, f"anchor_p count {src.count(anchor_p)}"
new_p = '    },\n' + projects_block + '\n  };\n\n  /* ============ PROJECT TEAMS'
src = src.replace(anchor_p, new_p)

# 2) Inject PROJECT_TEAMS entries after 2034 entry line.
# The 2034 team line ends with `</div>", "jobInfoUnits": 128}\n  };`
anchor_t = '"jobInfoUnits": 128}\n  };\n  Object.keys(PROJECT_TEAMS)'
assert src.count(anchor_t) == 1
new_t = '"jobInfoUnits": 128},\n' + teams_block + '\n  };\n  Object.keys(PROJECT_TEAMS)'
src = src.replace(anchor_t, new_t)

# 3) Inject extended arrays after PROJECTS['2034'].bva line.
# Find the last line of 2034 bva, which is `...\"ON\"]];\n` followed by blank + `\n\n  // ============ CHANGE LOG`
# The 2034 bva ends in `"ON"]];` at line 6874, before blank line.
# Simpler anchor: PROJECTS['2034'].bva = [...];\n\n\n  // ============ CHANGE LOG
# Use the last PROJECTS['2034'].bva as anchor (unique suffix)
import re
m = re.search(r"(PROJECTS\['2034'\]\.bva = \[.*?\];)", src)
assert m, 'could not find 2034 bva'
anchor_a = m.group(1)
new_a = anchor_a + '\n\n' + arrays_block
src = src.replace(anchor_a, new_a)

INDEX.write_text(src)
print('Injection complete. New line count:', src.count('\n'))
