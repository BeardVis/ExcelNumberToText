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
   IF(cellVal > 100000;
      CHOOSE(TRUNC(cellVal/100000); "сто"; "двісті"; "триста"; "чотириста"; "п'ятсот"; "шістсот"; "сімсот"; "вісімсот"; "дев'ятсот");
      ""
   )
   & IF(MOD(cellVal;100000) > 999;
      IF(TRUNC(MOD(cellVal;100000)/10000) = 1;
         IF(AND(cellVal >= 10000; cellVal < 20000); ""; " ") & CHOOSE(TRUNC(MOD(cellVal;10000)/1000)+1; "десять"; "одинадцять"; "дванадцять"; "тринадцять"; "чотирнадцять"; "п'ятнадцять"; "шістнадцять"; "сімнадцять"; "вісімнадцять"; "дев'ятнадцять") & " тисяч";
         IF(cellVal >= 100000; " "; "") & CHOOSE(TRUNC(MOD(cellVal;100000)/10000)+1;;; "двадцять"; "тридцять"; "сорок"; "п'ятдесят"; "шістдесят"; "сімдесят"; "вісімдесят"; "дев'яносто")
         & IF(INT(MOD(cellVal;10000))/1000 > 0;
            IF(cellVal > 9999; " "; "") & CHOOSE(TRUNC(MOD(cellVal;10000)/1000); "одна тисяча"; "дві тисячі"; "три тисячі"; "чотири тисячі"; "п'ять тисяч"; "шість тисяч"; "сім тисяч"; "вісім тисяч"; "дев'ять тисяч");
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
   & IF(AND(AND(MOD(cellVal;10) < 10; TRUNC(MOD(cellVal / 10; 10)) <> 1); MOD(TRUNC(cellVal);10) > 0);
      IF(cellVal < 10; ""; " ") & CHOOSE(TRUNC(MOD(cellVal;10) + 1); ; "одна"; "дві"; "три"; "чотири"; "п'ять"; "шість"; "сім"; "вісім"; "дев'ять") & " грив" & IF(TRUNC(MOD(cellVal;10))=0; "eнь"; IF(TRUNC(MOD(cellVal;10))=1; "ня"; IF(TRUNC(MOD(cellVal;10))<5; "ні"; "eнь")));
      IF(TRUNC(cellVal) = 0; "нуль"; "") & " гривень"
   )
)
& " " &
   IF(AND(MOD(INT(cellVal*100);100) < 20; MOD(INT(cellVal*100);100) > 9);
      CHOOSE(MOD(INT(cellVal*100);100)-9; "десять"; "одинадцять"; "дванадцять"; "тринадцять"; "чотирнадцять"; "п'ятнадцять"; "шістнадцять"; "сімнадцять"; "вісімнадцять"; "дев'ятнадцять") & " копійок";
      CHOOSE(IF(MOD(INT(cellVal*100);100) > 19; MOD(INT(cellVal*10);10) + 1; 1);;; "двадцять"; "тридцять"; "сорок"; "п'ятдесят"; "шістдесят"; "сімдесят"; "вісімдесят"; "дев'яносто")
      & IF(OR(MOD(INT(cellVal*100);100) < 10; MOD(INT(cellVal*100);10) < 10);
         IF(AND(MOD(INT(cellVal*100);100) > 20; MOD(INT(cellVal*100);10) <> 0); " "; "") & CHOOSE(IF(MOD(INT(cellVal*100);100) > 10; MOD(INT(cellVal*100);10)+1; MOD(INT(cellVal*100);100)+1); IF(MOD(INT(cellVal*100);100) = 0; "нуль"; ""); "одна"; "дві"; "три"; "чотири"; "п'ять"; "шість"; "сім"; "вісім"; "дев'ять") 
         & " копій" & IF(IF(MOD(INT(cellVal*100);100) > 10; MOD(INT(cellVal*100);10); MOD(INT(cellVal*100);100))=0; "ок"; IF(IF(MOD(INT(cellVal*100);100) > 10; MOD(INT(cellVal*100);10); MOD(INT(cellVal*100);100))=1; "ка"; IF(IF(MOD(INT(cellVal*100);100) > 10; MOD(INT(cellVal*100);10); MOD(INT(cellVal*100);100))<5; "ки"; "ок")));
         " копійок"
      )
   )
)
```


### Usage

Paste the formula into your Excel cell and replace `A1` in the second line with your cell reference.

### Note

The `LET` function is not supported in all Excel versions. Please check if your version supports it. If it doesn't, remove the `LET` function and replace the `cellVal` variable with your cell reference.
