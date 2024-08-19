# ExcelNumberToText

An Excel formula that converts numbers to text representing Ukrainian hryvnia currency.

The formula can convert positive numbers from `0` to `1,000,000`.

### Example:

|Number|Text|
|:---------|:---------------------------------------------------------------------------------|
|1,00|одна гривня 00 копійок|
|999999,99|дев'ятсот дев'яносто дев'ять тисяч дев'ятсот дев'яносто дев'ять гривень 99 копійок|

---

```VBScript
=
LET(cellVal; A1;
IF(cellVal < 1000000;
   IF(cellVal >= 100000;
      CHOOSE(TRUNC(cellVal / 100000); "сто"; "двісті"; "триста"; "чотириста"; "п'ятсот"; "шістсот"; "сімсот"; "вісімсот"; "дев'ятсот");
      ""
   )
   & IF(MOD(cellVal;100000) > 999;
      IF(TRUNC(MOD(cellVal;100000) / 10000) = 1;
         IF(AND(cellVal >= 10000; cellVal < 20000); ""; " ") & CHOOSE(TRUNC(MOD(cellVal;10000) / 1000) + 1; "десять"; "одинадцять"; "дванадцять"; "тринадцять"; "чотирнадцять"; "п'ятнадцять"; "шістнадцять"; "сімнадцять"; "вісімнадцять"; "дев'ятнадцять") & " тисяч";
         IF(AND(MOD(cellVal;100000) >= 19999; cellVal > 100000); " "; "") & CHOOSE(TRUNC(MOD(cellVal;100000) / 10000)+1;;; "двадцять"; "тридцять"; "сорок"; "п'ятдесят"; "шістдесят"; "сімдесят"; "вісімдесят"; "дев'яносто")
         & IF(INT(MOD(cellVal;10000)) / 1000 > 0;
            IF(cellVal > 9999; " "; "") & CHOOSE(TRUNC(MOD(cellVal;10000) / 1000); "одна тисяча"; "дві тисячі"; "три тисячі"; "чотири тисячі"; "п'ять тисяч"; "шість тисяч"; "сім тисяч"; "вісім тисяч"; "дев'ять тисяч");
            " тисяч"
         )
      );
      IF(AND(TRUNC(MOD(cellVal;10000)/1000) = 0; cellVal > 10000); " тисяч"; "")
   )
   & IF(MOD(cellVal;1000) > 99;
      IF(cellVal < 1000; ""; " ") & CHOOSE(TRUNC(MOD(cellVal;1000)/100) + 1;; "сто"; "двісті"; "триста"; "чотириста"; "п'ятсот"; "шістсот"; "сімсот"; "вісімсот"; "дев'ятсот");
      ""
   )
   & IF(MOD(cellVal;100) > 9;
      IF(TRUNC(MOD(cellVal;100)/10) = 1;
         IF(cellVal < 20; ""; " ") & CHOOSE(TRUNC(MOD(cellVal;10))+1; "десять"; "одинадцять"; "дванадцять"; "тринадцять"; "чотирнадцять"; "п'ятнадцять"; "шістнадцять"; "сімнадцять"; "вісімнадцять"; "дев'ятнадцять");
         IF(cellVal < 100; ""; " ") & CHOOSE(TRUNC(MOD(cellVal;100)/10)+1;;; "двадцять"; "тридцять"; "сорок"; "п'ятдесят"; "шістдесят"; "сімдесят"; "вісімдесят"; "дев'яносто")
      );
      ""
   )
   & IF(AND(MOD(cellVal;10) < 10; TRUNC(MOD(cellVal / 10; 10)) <> 1; MOD(TRUNC(cellVal);10) > 0);
      IF(cellVal < 10; ""; " ") & CHOOSE(TRUNC(MOD(cellVal;10) + 1); ; "одна гривня"; "дві гривні"; "три гривні"; "чотири гривні"; "п'ять гривень"; "шість гривень"; "сім гривень"; "вісім гривень"; "дев'ять гривень");
      IF(TRUNC(cellVal) = 0; "нуль"; "") & " гривень"
   )
)
&
IF(MOD(TRUNC(cellVal * 100);100) > 9; " "; " 0" )
& MOD(TRUNC(cellVal * 100);100)
& " копій" & IFS(
      OR(MOD(TRUNC(cellVal * 100);10) > 4; MOD(TRUNC(cellVal * 100);10) = 0; AND(MOD(TRUNC(cellVal * 100);100) >= 10; MOD(TRUNC(cellVal * 100);100) <= 20)); "ок";
      MOD(TRUNC(cellVal * 100);10) = 1; "ка";
      TRUE(); "ки"
)
)
```


### Usage

Paste the formula into your Excel cell and replace `A1` in the second line with your cell reference.

### Note

The `LET` function is not supported in all Excel versions. Please check if your version supports it. If it doesn't, remove the `LET` function and replace the `cellVal` variable with your cell reference.
