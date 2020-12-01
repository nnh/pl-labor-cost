#!/bin/env bash
sjis="charset=unknown-8bit"
temp=$(file -i ../programs/pl_labor_cost.bas |awk '{print $3}')
echo ${temp}
if [ $temp = $sjis ]; then
  iconv -f SHIFT-JIS -t UTF-8 ../programs/pl_labor_cost.bas > ../programs/temp.bas
  echo "SHIFT-JIS -> UTF-8"
else
  iconv -f UTF-8 -t SHIFT-JIS ../programs/pl_labor_cost.bas > ../programs/temp.bas
  echo "UTF-8 -> SHIFT-JIS"
fi
cp ../programs/temp.bas ../programs/pl_labor_cost.bas
rm ../programs/temp.bas
