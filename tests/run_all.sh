#!/bin/bash
# Full test suite: makes the test workbooks, exports each scenario with the tool's logic, then checks
# every export independently with verify.py (recalculates all formulas, re-derives every number).
# Run with: npm test   (needs node and python3 with openpyxl)
HERE="$(cd "$(dirname "$0")" && pwd)"; W="$HERE/out"; mkdir -p "$W"
node "$HERE/make_fixtures.js" "$W" || exit 1
pass=0; fail=0
run() { # name, env (VAR=value;VAR=value), source, verify args..., -- export picks...
  local name="$1" envs="$2" src="$3"; shift 3
  local vargs=() picks=(); while [ "$1" != "--" ] && [ $# -gt 0 ]; do vargs+=("$1"); shift; done; shift
  picks=("$@")
  local envarr=(); [ -n "$envs" ] && IFS=';' read -ra envarr <<< "$envs"
  if env "${envarr[@]}" node "$HERE/run_export.js" "$src" "$W/result.xlsx" "${picks[@]}" > "$W/log.txt" 2>&1 \
     && { if [[ "$name" == warn:* ]]; then grep -q 'not applied' "$W/log.txt"; else grep -q 'warnings \[\]' "$W/log.txt"; fi; } && python3 "$HERE/verify.py" "$src" "$W/result.xlsx" "${vargs[@]}" > "$W/check.txt" 2>&1; then
    pass=$((pass+1)); printf '  PASS  %s\n' "$name"
  else
    fail=$((fail+1)); printf '  FAIL  %s\n' "$name"; tail -8 "$W/log.txt" "$W/check.txt"
  fi
}
for SRC in "$W/formulas.xlsx" "$W/values_only.xlsx"; do
  echo "== $(basename "$SRC")"
  run "no changes"                         ""  "$SRC" --
  run "Product Fit + Range -> Others"      ""  "$SRC" "Product - Others=Product Fit+Product Range" -- "Product|Product Fit" "Product|Product Range"
  run "several groups, every Brand row"    ""  "$SRC" "Journey - Others=Ease & Convenience+Efficiency+Easy Logon" "Brand - Others=Brand Reputation+Trust & Security+Brand Image" "Distribution - Others=Queue Time" -- "Journey|Ease & Convenience" "Journey|Efficiency" "Journey|Easy Logon" "Brand|Brand Reputation" "Brand|Trust & Security" "Brand|Brand Image" "Distribution|Queue Time"
  run "combine Journey + Product"          "COMBINE=Journey,Product" "$SRC" --combine=Journey,Product --
  run "combine + Others across both"       "COMBINE=Journey,Product" "$SRC" --combine=Journey,Product "Journey & Product - Others=Efficiency+Product Fit" -- "Journey|Efficiency" "Product|Product Fit"
  run "combine non-adjacent Journey+Brand" "COMBINE=Journey,Brand" "$SRC" --combine=Journey,Brand "Distribution - Others=Accessibility+Queue Time" -- "Distribution|Queue Time" "Distribution|Accessibility"
  run "combine all four"                   "COMBINE=Journey,Product,Distribution,Brand" "$SRC" --combine=Journey,Product,Distribution,Brand --
  run "sort Pro-Det desc + Others"         "SORT=Bank A|2026 1H|net|desc" "$SRC" "Product - Others=Product Fit+Product Range" "--sort=Bank A|2026 1H|net|desc" -- "Product|Product Fit" "Product|Product Range"
  run "sort Rank asc in combined group"    "SORT=Bank F|2026 3Q YTD Interim|rank|asc;COMBINE=Journey,Product" "$SRC" --combine=Journey,Product "Journey & Product - Others=Efficiency+Product Fit" "--sort=Bank F|2026 3Q YTD Interim|rank|asc" -- "Journey|Efficiency" "Product|Product Fit"
  run "sort Rank chg desc"                 "SORT=Bank A|chg|rank|desc" "$SRC" "--sort=Bank A|chg|rank|desc" --
  echo "  -- named groupings"
  run "two named groupings in Journey"     'GROUPS=[{"name":"Digital","members":["Journey|Ease & Convenience","Journey|Efficiency","Journey|Easy Logon"]},{"name":"Service","members":["Journey|Informative","Journey|(Digital) Service Quality"]}]' "$SRC" "Digital=Ease & Convenience+Efficiency+Easy Logon" "Service=Informative+(Digital) Service Quality" --
  run "named + Others (bottom) in Product" 'GROUPS=[{"name":"Returns","members":["Product|Benefits & Rewards","Product|Level of Return"]}]' "$SRC" "Returns=Benefits & Rewards+Level of Return" "Product - Others=Product Fit+Product Range" -- "Product|Product Fit" "Product|Product Range"
  run "named groupings sorted, Others last" 'GROUPS=[{"name":"Digital","members":["Journey|Ease & Convenience","Journey|Efficiency"]},{"name":"Returns","members":["Product|Benefits & Rewards","Product|Level of Return"]}];SORT=Bank A|2026 1H|net|desc' "$SRC" "Digital=Ease & Convenience+Efficiency" "Returns=Benefits & Rewards+Level of Return" "Product - Others=Product Fit+Product Range" "--sort=Bank A|2026 1H|net|desc" -- "Product|Product Fit" "Product|Product Range"
  run "named non-Others kept at bottom"    'GROUPS=[{"name":"Tail","members":["Distribution|Accessibility","Distribution|Proactive Staff"],"pinLast":true}]' "$SRC" "Tail=Accessibility+Proactive Staff" --pinned=Tail --
  run "every Brand row in one grouping"    'GROUPS=[{"name":"Brand total","members":["Brand|Brand Reputation","Brand|Trust & Security","Brand|Brand Image"]}]' "$SRC" "Brand total=Brand Reputation+Trust & Security+Brand Image" --
  run "named combined group + mixed grouping" 'COMBINE=Journey,Product;COMBO_NAME=Customer Experience;GROUPS=[{"name":"Digital Mix","members":["Journey|Efficiency","Product|Product Fit"]}]' "$SRC" --combine=Journey,Product "--combo-name=Customer Experience" "Digital Mix=Efficiency+Product Fit" "Customer Experience - Others=Easy Logon+Product Range" -- "Journey|Easy Logon" "Product|Product Range"
  run "warn: mixed grouping, not combined" 'GROUPS=[{"name":"Digital Mix","members":["Journey|Efficiency","Product|Product Fit"]}]' "$SRC" --
done
echo "== five_banks.xlsx"
run "5 banks, new Brand rows, Others"      ""  "$W/five_banks.xlsx" "Product - Others=Product Fit+Product Range" "Brand - Others=Other Brand Images" -- "Product|Product Fit" "Product|Product Range" "Brand|Other Brand Images"
run "sort on a bank not in the file"       "SORT=Bank F|2026 1H|net|desc" "$W/five_banks.xlsx" --
echo "== blank_row.xlsx"
run "blank spacer row inside the table"    ""  "$W/blank_row.xlsx" "Brand - Others=Brand Image+Trust & Security" -- "Brand|Brand Image" "Brand|Trust & Security"
echo "== re-export of an export"
node "$HERE/run_export.js" "$W/formulas.xlsx" "$W/once.xlsx" "Product|Product Fit" "Product|Product Range" > /dev/null
run "export the exported file again"       ""  "$W/once.xlsx" --
echo "RESULT: $pass passed, $fail failed"
[ $fail -eq 0 ]
