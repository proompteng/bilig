# ODF 1.4 and Office Excel Function Coverage

## Source
- ODF 1.4 Spreadsheet OpenFormula Formula Functions
- OASIS OpenDocument v1.4 function requirements:
  - https://docs.oasis-open.org/office/OpenDocument/v1.4/os/v1.4-os.html
- Microsoft Office Excel functions by category:
  - https://support.microsoft.com/en-us/office/excel-functions-by-category-5f91f4e9-7b42-46d2-9bd1-63f26a86c0eb
- Scope: ODF mandatory functions + Office by-category function list (merged into one list).

## Coverage Summary
- ODF mandatory function count: **277**
- Office function count (cleaned scrape): **505**
- Overlap (present in both): **260**
- ODF-only (mandatory, not listed by Office): **17**
- Office-only (not in ODF 1.4 mandatory): **245**

## Full unified function list (ODF 1.4 mandatory + Office category)

| Function | ODF status | In Office list | Implemented in codebase |
| --- | --- | --- | --- |
| ABS | Not in ODF 1.4 mandatory | Yes | Yes |
| ACCRINT | Implemented | Yes | Yes |
| ACCRINTM | Implemented | Yes | Yes |
| ACOS | Not in ODF 1.4 mandatory | Yes | Yes |
| ACOSH | Implemented | Yes | Yes |
| ACOT | Implemented | Yes | Yes |
| ACOTH | Implemented | Yes | Yes |
| ADDRESS | Placeholder (blocked) | Yes | No |
| AGGREGATE | Not in ODF 1.4 mandatory | Yes | Yes |
| AMORDEGRC | Not in ODF 1.4 mandatory | Yes | Yes |
| AMORLINC | Implemented | Yes | Yes |
| AND | Not in ODF 1.4 mandatory | Yes | Yes |
| ARABIC | Implemented | Yes | Yes |
| AREAS | Placeholder (blocked) | Yes | No |
| ARRAYTOTEXT | Not in ODF 1.4 mandatory | Yes | No |
| ASC | Placeholder (blocked) | Yes | No |
| ASIN | Not in ODF 1.4 mandatory | Yes | Yes |
| ASINH | Implemented | Yes | Yes |
| ATAN | Not in ODF 1.4 mandatory | Yes | Yes |
| ATAN2 | Not in ODF 1.4 mandatory | Yes | Yes |
| ATANH | Implemented | Yes | Yes |
| AVEDEV | Placeholder (blocked) | Yes | No |
| AVERAGE | Not in ODF 1.4 mandatory | Yes | Yes |
| AVERAGEA | Placeholder (blocked) | Yes | No |
| AVERAGEIF | Not in ODF 1.4 mandatory | Yes | Yes |
| AVERAGEIFS | Implemented | Yes | Yes |
| BAHTTEXT | Not in ODF 1.4 mandatory | Yes | No |
| BASE | Implemented | Yes | Yes |
| BESSELI | Placeholder (blocked) | Yes | No |
| BESSELJ | Placeholder (blocked) | Yes | No |
| BESSELK | Placeholder (blocked) | Yes | No |
| BESSELY | Placeholder (blocked) | Yes | No |
| BETA.DIST | Not in ODF 1.4 mandatory | Yes | No |
| BETA.INV | Not in ODF 1.4 mandatory | Yes | No |
| BETADIST | Placeholder (blocked) | Yes | No |
| BETAINV | Placeholder (blocked) | Yes | No |
| BIN2DEC | Placeholder (blocked) | Yes | No |
| BIN2HEX | Placeholder (blocked) | Yes | No |
| BIN2OCT | Placeholder (blocked) | Yes | No |
| BINOM.DIST | Not in ODF 1.4 mandatory | Yes | No |
| BINOM.DIST.RANGE | Placeholder (blocked) | Yes | No |
| BINOM.INV | Not in ODF 1.4 mandatory | Yes | No |
| BINOMDIST | Placeholder (blocked) | Yes | No |
| BITAND | Placeholder (blocked) | Yes | No |
| BITLSHIFT | Placeholder (blocked) | Yes | No |
| BITOR | Placeholder (blocked) | Yes | No |
| BITRSHIFT | Placeholder (blocked) | Yes | No |
| BITXOR | Placeholder (blocked) | Yes | No |
| BYCOL | Not in ODF 1.4 mandatory | Yes | Yes |
| BYROW | Not in ODF 1.4 mandatory | Yes | Yes |
| CALL | Not in ODF 1.4 mandatory | Yes | No |
| CEILING | Implemented | Yes | Yes |
| CEILING.MATH | Not in ODF 1.4 mandatory | Yes | Yes |
| CEILING.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| CELL | Not in ODF 1.4 mandatory | Yes | No |
| CHAR | Placeholder (blocked) | Yes | No |
| CHIDIST | Not in ODF 1.4 mandatory | Yes | No |
| CHIINV | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.DIST | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.DIST.RT | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.INV | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.INV.RT | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.TEST | Not in ODF 1.4 mandatory | Yes | No |
| CHISQDIST | Missing | No | No |
| CHISQINV | Missing | No | No |
| CHITEST | Not in ODF 1.4 mandatory | Yes | No |
| CHOOSE | Not in ODF 1.4 mandatory | Yes | No |
| CHOOSECOLS | Not in ODF 1.4 mandatory | Yes | No |
| CHOOSEROWS | Not in ODF 1.4 mandatory | Yes | No |
| CLEAN | Placeholder (blocked) | Yes | No |
| CODE | Placeholder (blocked) | Yes | No |
| COLUMN | Placeholder (blocked) | Yes | No |
| COLUMNS | Not in ODF 1.4 mandatory | Yes | No |
| COMBIN | Implemented | Yes | Yes |
| COMBINA | Implemented | Yes | Yes |
| COMPLEX | Placeholder (blocked) | Yes | No |
| CONCAT | Not in ODF 1.4 mandatory | Yes | Yes |
| CONCATENATE | Placeholder (blocked) | Yes | No |
| CONFIDENCE | Placeholder (blocked) | Yes | No |
| CONFIDENCE.NORM | Not in ODF 1.4 mandatory | Yes | No |
| CONFIDENCE.T | Not in ODF 1.4 mandatory | Yes | No |
| CONVERT | Placeholder (blocked) | Yes | No |
| CORREL | Placeholder (blocked) | Yes | No |
| COS | Not in ODF 1.4 mandatory | Yes | Yes |
| COSH | Implemented | Yes | Yes |
| COT | Implemented | Yes | Yes |
| COTH | Implemented | Yes | Yes |
| COUNT | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTA | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTBLANK | Not in ODF 1.4 mandatory | Yes | No |
| COUNTIF | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTIFS | Implemented | Yes | Yes |
| COUPDAYBS | Placeholder (blocked) | Yes | No |
| COUPDAYS | Placeholder (blocked) | Yes | No |
| COUPDAYSNC | Placeholder (blocked) | Yes | No |
| COUPNCD | Placeholder (blocked) | Yes | No |
| COUPNUM | Placeholder (blocked) | Yes | No |
| COUPPCD | Placeholder (blocked) | Yes | No |
| COVAR | Placeholder (blocked) | Yes | No |
| COVARIANCE.P | Not in ODF 1.4 mandatory | Yes | No |
| COVARIANCE.S | Not in ODF 1.4 mandatory | Yes | No |
| CRITBINOM | Placeholder (blocked) | Yes | No |
| CSC | Implemented | Yes | Yes |
| CSCH | Implemented | Yes | Yes |
| CUBEKPIMEMBER | Not in ODF 1.4 mandatory | Yes | No |
| CUBEMEMBER | Not in ODF 1.4 mandatory | Yes | No |
| CUBEMEMBERPROPERTY | Not in ODF 1.4 mandatory | Yes | No |
| CUBERANKEDMEMBER | Not in ODF 1.4 mandatory | Yes | No |
| CUBESET | Not in ODF 1.4 mandatory | Yes | No |
| CUBESETCOUNT | Not in ODF 1.4 mandatory | Yes | No |
| CUBEVALUE | Not in ODF 1.4 mandatory | Yes | No |
| CUMIPMT | Placeholder (blocked) | Yes | No |
| CUMPRINC | Placeholder (blocked) | Yes | No |
| DATE | Not in ODF 1.4 mandatory | Yes | Yes |
| DATEDIF | Placeholder (blocked) | Yes | No |
| DATEVALUE | Placeholder (blocked) | Yes | No |
| DAVERAGE | Not in ODF 1.4 mandatory | Yes | No |
| DAY | Not in ODF 1.4 mandatory | Yes | Yes |
| DAYS | Placeholder (blocked) | Yes | No |
| DAYS360 | Placeholder (blocked) | Yes | No |
| DB | Placeholder (blocked) | Yes | No |
| DBCS | Not in ODF 1.4 mandatory | Yes | No |
| DCOUNT | Not in ODF 1.4 mandatory | Yes | No |
| DCOUNTA | Not in ODF 1.4 mandatory | Yes | No |
| DDB | Not in ODF 1.4 mandatory | Yes | No |
| DDE | Missing | No | No |
| DEC2BIN | Placeholder (blocked) | Yes | No |
| DEC2HEX | Placeholder (blocked) | Yes | No |
| DEC2OCT | Placeholder (blocked) | Yes | No |
| DECIMAL | Implemented | Yes | Yes |
| DEGREES | Not in ODF 1.4 mandatory | Yes | Yes |
| DELTA | Placeholder (blocked) | Yes | No |
| DETECTLANGUAGE | Not in ODF 1.4 mandatory | Yes | No |
| DEVSQ | Placeholder (blocked) | Yes | No |
| DGET | Not in ODF 1.4 mandatory | Yes | No |
| DISC | Placeholder (blocked) | Yes | No |
| DMAX | Not in ODF 1.4 mandatory | Yes | No |
| DMIN | Not in ODF 1.4 mandatory | Yes | No |
| DOLLAR | Not in ODF 1.4 mandatory | Yes | No |
| DOLLARDE | Placeholder (blocked) | Yes | No |
| DOLLARFR | Placeholder (blocked) | Yes | No |
| DPRODUCT | Not in ODF 1.4 mandatory | Yes | No |
| DROP | Not in ODF 1.4 mandatory | Yes | No |
| DSTDEV | Not in ODF 1.4 mandatory | Yes | No |
| DSTDEVP | Not in ODF 1.4 mandatory | Yes | No |
| DSUM | Not in ODF 1.4 mandatory | Yes | No |
| DURATION | Placeholder (blocked) | Yes | No |
| DVAR | Not in ODF 1.4 mandatory | Yes | No |
| DVARP | Not in ODF 1.4 mandatory | Yes | No |
| EDATE | Implemented | Yes | Yes |
| EFFECT | Placeholder (blocked) | Yes | No |
| ENCODEURL | Not in ODF 1.4 mandatory | Yes | No |
| EOMONTH | Implemented | Yes | Yes |
| ERF | Placeholder (blocked) | Yes | No |
| ERF.PRECISE | Not in ODF 1.4 mandatory | Yes | No |
| ERFC | Placeholder (blocked) | Yes | No |
| ERFC.PRECISE | Not in ODF 1.4 mandatory | Yes | No |
| ERROR.TYPE | Placeholder (blocked) | Yes | No |
| EUROCONVERT | Placeholder (blocked) | Yes | No |
| EVEN | Not in ODF 1.4 mandatory | Yes | Yes |
| EXACT | Not in ODF 1.4 mandatory | Yes | Yes |
| EXP | Not in ODF 1.4 mandatory | Yes | Yes |
| EXPAND | Not in ODF 1.4 mandatory | Yes | No |
| EXPON.DIST | Not in ODF 1.4 mandatory | Yes | No |
| EXPONDIST | Placeholder (blocked) | Yes | No |
| F.DIST | Not in ODF 1.4 mandatory | Yes | No |
| F.DIST.RT | Not in ODF 1.4 mandatory | Yes | No |
| F.INV | Not in ODF 1.4 mandatory | Yes | No |
| F.INV.RT | Not in ODF 1.4 mandatory | Yes | No |
| F.TEST | Not in ODF 1.4 mandatory | Yes | No |
| FACT | Not in ODF 1.4 mandatory | Yes | Yes |
| FACTDOUBLE | Implemented | Yes | Yes |
| FALSE | Not in ODF 1.4 mandatory | Yes | Yes |
| FDIST | Placeholder (blocked) | Yes | No |
| FILTER | Not in ODF 1.4 mandatory | Yes | Yes |
| FILTERXML | Not in ODF 1.4 mandatory | Yes | No |
| FIND | Not in ODF 1.4 mandatory | Yes | Yes |
| FINDB | Missing | Yes | No |
| FINV | Placeholder (blocked) | Yes | No |
| FISHER | Placeholder (blocked) | Yes | No |
| FISHERINV | Placeholder (blocked) | Yes | No |
| FIXED | Placeholder (blocked) | Yes | No |
| FLOOR | Implemented | Yes | Yes |
| FLOOR.MATH | Not in ODF 1.4 mandatory | Yes | Yes |
| FLOOR.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| FORECAST | Placeholder (blocked) | No | No |
| FORMULA | Missing | No | No |
| FORMULATEXT | Not in ODF 1.4 mandatory | Yes | No |
| FREQUENCY | Placeholder (blocked) | Yes | No |
| FTEST | Placeholder (blocked) | Yes | No |
| FV | Not in ODF 1.4 mandatory | Yes | No |
| FVSCHEDULE | Placeholder (blocked) | Yes | No |
| GAMMA | Placeholder (blocked) | Yes | No |
| GAMMA.DIST | Not in ODF 1.4 mandatory | Yes | No |
| GAMMA.INV | Not in ODF 1.4 mandatory | Yes | No |
| GAMMADIST | Placeholder (blocked) | Yes | No |
| GAMMAINV | Placeholder (blocked) | Yes | No |
| GAMMALN | Placeholder (blocked) | Yes | No |
| GAMMALN.PRECISE | Not in ODF 1.4 mandatory | Yes | No |
| GAUSS | Placeholder (blocked) | Yes | No |
| GCD | Implemented | Yes | Yes |
| GEOMEAN | Placeholder (blocked) | Yes | No |
| GESTEP | Placeholder (blocked) | Yes | No |
| GETPIVOTDATA | Placeholder (blocked) | Yes | No |
| GROUPBY | Not in ODF 1.4 mandatory | Yes | No |
| GROWTH | Placeholder (blocked) | Yes | No |
| HARMEAN | Placeholder (blocked) | Yes | No |
| HEX2BIN | Placeholder (blocked) | Yes | No |
| HEX2DEC | Placeholder (blocked) | Yes | No |
| HEX2OCT | Placeholder (blocked) | Yes | No |
| HLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| HOUR | Not in ODF 1.4 mandatory | Yes | Yes |
| HSTACK | Not in ODF 1.4 mandatory | Yes | No |
| HYPERLINK | Placeholder (blocked) | Yes | No |
| HYPGEOM.DIST | Not in ODF 1.4 mandatory | Yes | No |
| HYPGEOMDIST | Placeholder (blocked) | Yes | No |
| IF | Not in ODF 1.4 mandatory | Yes | Yes |
| IFERROR | Implemented | Yes | Yes |
| IFNA | Implemented | Yes | Yes |
| IFS | Not in ODF 1.4 mandatory | Yes | Yes |
| IMABS | Placeholder (blocked) | Yes | No |
| IMAGE | Not in ODF 1.4 mandatory | Yes | No |
| IMAGINARY | Placeholder (blocked) | Yes | No |
| IMARGUMENT | Placeholder (blocked) | Yes | No |
| IMCONJUGATE | Placeholder (blocked) | Yes | No |
| IMCOS | Placeholder (blocked) | Yes | No |
| IMCOSH | Not in ODF 1.4 mandatory | Yes | No |
| IMCOT | Placeholder (blocked) | Yes | No |
| IMCSC | Placeholder (blocked) | Yes | No |
| IMCSCH | Placeholder (blocked) | Yes | No |
| IMDIV | Placeholder (blocked) | Yes | No |
| IMEXP | Placeholder (blocked) | Yes | No |
| IMLN | Placeholder (blocked) | Yes | No |
| IMLOG10 | Placeholder (blocked) | Yes | No |
| IMLOG2 | Placeholder (blocked) | Yes | No |
| IMPOWER | Placeholder (blocked) | Yes | No |
| IMPRODUCT | Placeholder (blocked) | Yes | No |
| IMREAL | Placeholder (blocked) | Yes | No |
| IMSEC | Placeholder (blocked) | Yes | No |
| IMSECH | Placeholder (blocked) | Yes | No |
| IMSIN | Placeholder (blocked) | Yes | No |
| IMSINH | Not in ODF 1.4 mandatory | Yes | No |
| IMSQRT | Placeholder (blocked) | Yes | No |
| IMSUB | Placeholder (blocked) | Yes | No |
| IMSUM | Placeholder (blocked) | Yes | No |
| IMTAN | Placeholder (blocked) | Yes | No |
| INDEX | Not in ODF 1.4 mandatory | Yes | Yes |
| INDIRECT | Placeholder (blocked) | Yes | No |
| INFO | Placeholder (blocked) | Yes | No |
| INT | Not in ODF 1.4 mandatory | Yes | Yes |
| INTERCEPT | Placeholder (blocked) | Yes | No |
| INTRATE | Placeholder (blocked) | Yes | No |
| IPMT | Placeholder (blocked) | Yes | No |
| IRR | Not in ODF 1.4 mandatory | Yes | No |
| ISEVEN | Placeholder (blocked) | Yes | No |
| ISFORMULA | Placeholder (blocked) | Yes | No |
| ISO.CEILING | Not in ODF 1.4 mandatory | Yes | Yes |
| ISODD | Placeholder (blocked) | No | No |
| ISOMITTED | Not in ODF 1.4 mandatory | Yes | No |
| ISOWEEKNUM | Placeholder (blocked) | Yes | No |
| ISPMT | Placeholder (blocked) | Yes | No |
| ISREF | Placeholder (blocked) | No | No |
| JIS | Placeholder (blocked) | No | No |
| KURT | Placeholder (blocked) | Yes | No |
| LAMBDA | Not in ODF 1.4 mandatory | Yes | Yes |
| LARGE | Placeholder (blocked) | Yes | No |
| LCM | Implemented | Yes | Yes |
| LEFT | Not in ODF 1.4 mandatory | Yes | Yes |
| LEFTB | Missing | Yes | No |
| LEGACY.CHIDIST | Legacy not implemented | No | No |
| LEGACY.CHIINV | Legacy not implemented | No | No |
| LEGACY.CHITEST | Legacy not implemented | No | No |
| LEGACY.FDIST | Legacy not implemented | No | No |
| LEGACY.FINV | Legacy not implemented | No | No |
| LEGACY.NORMSDIST | Legacy not implemented | No | No |
| LEGACY.NORMSINV | Legacy not implemented | No | No |
| LEN | Not in ODF 1.4 mandatory | Yes | Yes |
| LENB | Missing | Yes | No |
| LET | Not in ODF 1.4 mandatory | Yes | Yes |
| LINEST | Placeholder (blocked) | Yes | No |
| LN | Not in ODF 1.4 mandatory | Yes | Yes |
| LOG | Not in ODF 1.4 mandatory | Yes | Yes |
| LOG10 | Not in ODF 1.4 mandatory | Yes | Yes |
| LOGEST | Placeholder (blocked) | Yes | No |
| LOGINV | Placeholder (blocked) | Yes | No |
| LOGNORM.DIST | Not in ODF 1.4 mandatory | Yes | No |
| LOGNORM.INV | Not in ODF 1.4 mandatory | Yes | No |
| LOGNORMDIST | Placeholder (blocked) | Yes | No |
| LOOKUP | Placeholder (blocked) | Yes | No |
| LOWER | Not in ODF 1.4 mandatory | Yes | Yes |
| MAKEARRAY | Not in ODF 1.4 mandatory | Yes | Yes |
| MAP | Not in ODF 1.4 mandatory | Yes | Yes |
| MATCH | Not in ODF 1.4 mandatory | Yes | Yes |
| MAX | Not in ODF 1.4 mandatory | Yes | Yes |
| MAXA | Placeholder (blocked) | Yes | No |
| MAXIFS | Not in ODF 1.4 mandatory | Yes | No |
| MDETERM | Implemented | Yes | Yes |
| MDURATION | Placeholder (blocked) | Yes | No |
| MEDIAN | Placeholder (blocked) | Yes | No |
| MID | Not in ODF 1.4 mandatory | Yes | Yes |
| MIDB | Missing | Yes | No |
| MIN | Not in ODF 1.4 mandatory | Yes | Yes |
| MINA | Placeholder (blocked) | Yes | No |
| MINIFS | Not in ODF 1.4 mandatory | Yes | No |
| MINUTE | Not in ODF 1.4 mandatory | Yes | Yes |
| MINVERSE | Implemented | Yes | Yes |
| MIRR | Placeholder (blocked) | Yes | No |
| MMULT | Implemented | Yes | Yes |
| MOD | Not in ODF 1.4 mandatory | Yes | Yes |
| MODE | Placeholder (blocked) | Yes | No |
| MODE.MULT | Not in ODF 1.4 mandatory | Yes | No |
| MODE.SNGL | Not in ODF 1.4 mandatory | Yes | No |
| MONTH | Not in ODF 1.4 mandatory | Yes | Yes |
| MROUND | Implemented | Yes | Yes |
| MULTINOMIAL | Implemented | Yes | Yes |
| MULTIPLE.OPERATIONS | Missing | No | No |
| MUNIT | Implemented | Yes | Yes |
| N | Not in ODF 1.4 mandatory | Yes | No |
| NA | Not in ODF 1.4 mandatory | Yes | Yes |
| NEGBINOM.DIST | Not in ODF 1.4 mandatory | Yes | No |
| NEGBINOMDIST | Placeholder (blocked) | Yes | No |
| NETWORKDAYS | Placeholder (blocked) | Yes | No |
| NETWORKDAYS.INTL | Not in ODF 1.4 mandatory | Yes | No |
| NOMINAL | Placeholder (blocked) | Yes | No |
| NORM.DIST | Not in ODF 1.4 mandatory | Yes | No |
| NORM.INV | Not in ODF 1.4 mandatory | Yes | No |
| NORM.S.DIST | Not in ODF 1.4 mandatory | Yes | No |
| NORM.S.INV | Not in ODF 1.4 mandatory | Yes | No |
| NORMDIST | Placeholder (blocked) | Yes | No |
| NORMINV | Placeholder (blocked) | Yes | No |
| NORMSDIST | Not in ODF 1.4 mandatory | Yes | No |
| NORMSINV | Not in ODF 1.4 mandatory | Yes | No |
| NOT | Not in ODF 1.4 mandatory | Yes | Yes |
| NOW | Not in ODF 1.4 mandatory | Yes | Yes |
| NPER | Not in ODF 1.4 mandatory | Yes | No |
| NPV | Not in ODF 1.4 mandatory | Yes | No |
| NUMBERVALUE | Placeholder (blocked) | Yes | No |
| OCT2BIN | Placeholder (blocked) | Yes | No |
| OCT2DEC | Placeholder (blocked) | Yes | No |
| OCT2HEX | Placeholder (blocked) | Yes | No |
| ODD | Not in ODF 1.4 mandatory | Yes | Yes |
| ODDFPRICE | Placeholder (blocked) | Yes | No |
| ODDFYIELD | Placeholder (blocked) | Yes | No |
| ODDLPRICE | Placeholder (blocked) | Yes | No |
| ODDLYIELD | Placeholder (blocked) | Yes | No |
| OFFSET | Placeholder (blocked) | Yes | No |
| OR | Not in ODF 1.4 mandatory | Yes | Yes |
| PDURATION | Placeholder (blocked) | Yes | No |
| PEARSON | Placeholder (blocked) | Yes | No |
| PERCENTILE | Placeholder (blocked) | Yes | No |
| PERCENTILE.EXC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTILE.INC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTOF | Not in ODF 1.4 mandatory | Yes | Yes |
| PERCENTRANK | Placeholder (blocked) | Yes | No |
| PERCENTRANK.EXC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTRANK.INC | Not in ODF 1.4 mandatory | Yes | No |
| PERMUT | Placeholder (blocked) | Yes | No |
| PERMUTATIONA | Placeholder (blocked) | Yes | No |
| PHI | Placeholder (blocked) | Yes | No |
| PHONETIC | Not in ODF 1.4 mandatory | Yes | No |
| PI | Not in ODF 1.4 mandatory | Yes | Yes |
| PIVOTBY | Not in ODF 1.4 mandatory | Yes | No |
| PMT | Not in ODF 1.4 mandatory | Yes | No |
| POISSON | Placeholder (blocked) | Yes | No |
| POISSON.DIST | Not in ODF 1.4 mandatory | Yes | No |
| POWER | Not in ODF 1.4 mandatory | Yes | Yes |
| PPMT | Placeholder (blocked) | Yes | No |
| PRICE | Placeholder (blocked) | Yes | No |
| PRICEDISC | Placeholder (blocked) | Yes | No |
| PRICEMAT | Placeholder (blocked) | Yes | No |
| PROB | Placeholder (blocked) | Yes | No |
| PRODUCT | Not in ODF 1.4 mandatory | Yes | Yes |
| PROPER | Not in ODF 1.4 mandatory | Yes | No |
| PV | Not in ODF 1.4 mandatory | Yes | No |
| QUARTILE | Placeholder (blocked) | Yes | No |
| QUARTILE.EXC | Not in ODF 1.4 mandatory | Yes | No |
| QUARTILE.INC | Not in ODF 1.4 mandatory | Yes | No |
| QUOTIENT | Implemented | Yes | Yes |
| RADIANS | Not in ODF 1.4 mandatory | Yes | Yes |
| RAND | Implemented | Yes | Yes |
| RANDARRAY | Not in ODF 1.4 mandatory | Yes | Yes |
| RANDBETWEEN | Implemented | Yes | Yes |
| RANK | Placeholder (blocked) | Yes | No |
| RANK.AVG | Not in ODF 1.4 mandatory | Yes | No |
| RANK.EQ | Not in ODF 1.4 mandatory | Yes | No |
| RATE | Not in ODF 1.4 mandatory | Yes | No |
| RECEIVED | Placeholder (blocked) | Yes | No |
| REDUCE | Not in ODF 1.4 mandatory | Yes | Yes |
| REGEXEXTRACT | Not in ODF 1.4 mandatory | Yes | No |
| REGEXREPLACE | Not in ODF 1.4 mandatory | Yes | No |
| REGEXTEST | Not in ODF 1.4 mandatory | Yes | No |
| REGISTER.ID | Not in ODF 1.4 mandatory | Yes | No |
| REPLACE | Not in ODF 1.4 mandatory | Yes | No |
| REPLACEB | Missing | Yes | No |
| REPT | Not in ODF 1.4 mandatory | Yes | No |
| RIGHT | Not in ODF 1.4 mandatory | Yes | Yes |
| RIGHTB | Missing | Yes | No |
| ROMAN | Implemented | Yes | Yes |
| ROUND | Not in ODF 1.4 mandatory | Yes | Yes |
| ROUNDDOWN | Implemented | Yes | Yes |
| ROUNDUP | Implemented | Yes | Yes |
| ROW | Placeholder (blocked) | Yes | No |
| ROWS | Not in ODF 1.4 mandatory | Yes | No |
| RRI | Placeholder (blocked) | Yes | No |
| RSQ | Placeholder (blocked) | Yes | No |
| RTD | Not in ODF 1.4 mandatory | Yes | No |
| SCAN | Not in ODF 1.4 mandatory | Yes | Yes |
| SEARCH | Implemented | Yes | Yes |
| SEARCHB | Missing | Yes | No |
| SEC | Implemented | Yes | Yes |
| SECH | Implemented | Yes | Yes |
| SECOND | Not in ODF 1.4 mandatory | Yes | Yes |
| SEQUENCE | Not in ODF 1.4 mandatory | Yes | Yes |
| SERIESSUM | Implemented | Yes | Yes |
| SHEET | Placeholder (blocked) | Yes | No |
| SHEETS | Placeholder (blocked) | Yes | No |
| SIGN | Implemented | Yes | Yes |
| SIN | Not in ODF 1.4 mandatory | Yes | Yes |
| SINH | Implemented | Yes | Yes |
| SKEW | Placeholder (blocked) | Yes | No |
| SKEW.P | Not in ODF 1.4 mandatory | Yes | No |
| SKEWP | Missing | No | No |
| SLN | Not in ODF 1.4 mandatory | Yes | No |
| SLOPE | Placeholder (blocked) | Yes | No |
| SMALL | Placeholder (blocked) | Yes | No |
| SORT | Not in ODF 1.4 mandatory | Yes | No |
| SORTBY | Not in ODF 1.4 mandatory | Yes | No |
| SQRT | Not in ODF 1.4 mandatory | Yes | Yes |
| SQRTPI | Implemented | Yes | Yes |
| STANDARDIZE | Placeholder (blocked) | Yes | No |
| STDEV | Not in ODF 1.4 mandatory | Yes | No |
| STDEV.P | Not in ODF 1.4 mandatory | Yes | No |
| STDEV.S | Not in ODF 1.4 mandatory | Yes | No |
| STDEVA | Placeholder (blocked) | Yes | No |
| STDEVP | Not in ODF 1.4 mandatory | Yes | No |
| STDEVPA | Placeholder (blocked) | Yes | No |
| STEYX | Placeholder (blocked) | Yes | No |
| STOCKHISTORY | Not in ODF 1.4 mandatory | Yes | No |
| SUBSTITUTE | Not in ODF 1.4 mandatory | Yes | No |
| SUBTOTAL | Implemented | Yes | Yes |
| SUM | Not in ODF 1.4 mandatory | Yes | Yes |
| SUMIF | Not in ODF 1.4 mandatory | Yes | Yes |
| SUMIFS | Implemented | Yes | Yes |
| SUMPRODUCT | Implemented | Yes | Yes |
| SUMSQ | Implemented | Yes | Yes |
| SUMX2MY2 | Implemented | Yes | Yes |
| SUMX2PY2 | Implemented | Yes | Yes |
| SUMXMY2 | Implemented | Yes | Yes |
| SWITCH | Not in ODF 1.4 mandatory | Yes | Yes |
| SYD | Not in ODF 1.4 mandatory | Yes | No |
| T | Not in ODF 1.4 mandatory | Yes | No |
| T.DIST | Not in ODF 1.4 mandatory | Yes | No |
| T.DIST.2T | Not in ODF 1.4 mandatory | Yes | No |
| T.DIST.RT | Not in ODF 1.4 mandatory | Yes | No |
| T.INV | Not in ODF 1.4 mandatory | Yes | No |
| T.INV.2T | Not in ODF 1.4 mandatory | Yes | No |
| T.TEST | Not in ODF 1.4 mandatory | Yes | No |
| TAKE | Not in ODF 1.4 mandatory | Yes | No |
| TAN | Not in ODF 1.4 mandatory | Yes | Yes |
| TANH | Implemented | Yes | Yes |
| TBILLEQ | Placeholder (blocked) | Yes | No |
| TBILLPRICE | Placeholder (blocked) | Yes | No |
| TBILLYIELD | Placeholder (blocked) | Yes | No |
| TDIST | Not in ODF 1.4 mandatory | Yes | No |
| TEXT | Placeholder (blocked) | Yes | No |
| TEXTAFTER | Not in ODF 1.4 mandatory | Yes | No |
| TEXTBEFORE | Not in ODF 1.4 mandatory | Yes | Yes |
| TEXTJOIN | Not in ODF 1.4 mandatory | Yes | No |
| TEXTSPLIT | Not in ODF 1.4 mandatory | Yes | No |
| TIME | Not in ODF 1.4 mandatory | Yes | Yes |
| TIMEVALUE | Placeholder (blocked) | Yes | No |
| TINV | Placeholder (blocked) | Yes | No |
| TOCOL | Not in ODF 1.4 mandatory | Yes | No |
| TODAY | Not in ODF 1.4 mandatory | Yes | Yes |
| TOROW | Not in ODF 1.4 mandatory | Yes | No |
| TRANSLATE | Not in ODF 1.4 mandatory | Yes | No |
| TRANSPOSE | Placeholder (blocked) | Yes | No |
| TREND | Placeholder (blocked) | Yes | No |
| TRIM | Not in ODF 1.4 mandatory | Yes | Yes |
| TRIMMEAN | Placeholder (blocked) | Yes | No |
| TRIMRANGE | Not in ODF 1.4 mandatory | Yes | No |
| TRUE | Not in ODF 1.4 mandatory | Yes | Yes |
| TRUNC | Not in ODF 1.4 mandatory | Yes | Yes |
| TTEST | Placeholder (blocked) | Yes | No |
| TYPE | Placeholder (blocked) | Yes | No |
| UNICHAR | Placeholder (blocked) | Yes | No |
| UNICODE | Placeholder (blocked) | Yes | No |
| UNIQUE | Not in ODF 1.4 mandatory | Yes | Yes |
| UPPER | Not in ODF 1.4 mandatory | Yes | Yes |
| USE.THE.COUNTIF | Not in ODF 1.4 mandatory | Yes | No |
| VALUE | Not in ODF 1.4 mandatory | Yes | Yes |
| VALUETOTEXT | Not in ODF 1.4 mandatory | Yes | No |
| VAR | Not in ODF 1.4 mandatory | Yes | No |
| VAR.P | Not in ODF 1.4 mandatory | Yes | No |
| VAR.S | Not in ODF 1.4 mandatory | Yes | No |
| VARA | Placeholder (blocked) | Yes | No |
| VARP | Not in ODF 1.4 mandatory | Yes | No |
| VARPA | Placeholder (blocked) | Yes | No |
| VDB | Placeholder (blocked) | Yes | No |
| VLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| VSTACK | Not in ODF 1.4 mandatory | Yes | No |
| WEBSERVICE | Not in ODF 1.4 mandatory | Yes | No |
| WEEKDAY | Not in ODF 1.4 mandatory | Yes | Yes |
| WEEKNUM | Placeholder (blocked) | Yes | No |
| WEIBULL | Placeholder (blocked) | Yes | No |
| WEIBULL.DIST | Not in ODF 1.4 mandatory | Yes | No |
| WORKDAY | Placeholder (blocked) | Yes | No |
| WORKDAY.INTL | Not in ODF 1.4 mandatory | Yes | No |
| WRAPCOLS | Not in ODF 1.4 mandatory | Yes | No |
| WRAPROWS | Not in ODF 1.4 mandatory | Yes | No |
| XIRR | Placeholder (blocked) | Yes | No |
| XLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| XMATCH | Not in ODF 1.4 mandatory | Yes | Yes |
| XNPV | Placeholder (blocked) | Yes | No |
| XOR | Implemented | Yes | Yes |
| YEAR | Not in ODF 1.4 mandatory | Yes | Yes |
| YEARFRAC | Placeholder (blocked) | Yes | No |
| YIELD | Placeholder (blocked) | Yes | No |
| YIELDDISC | Placeholder (blocked) | Yes | No |
| YIELDMAT | Placeholder (blocked) | Yes | No |
| Z.TEST | Not in ODF 1.4 mandatory | Yes | No |
| ZTEST | Placeholder (blocked) | Yes | No |
