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

## Current code coverage snapshot
- Implemented in codebase: **312**
- Not yet implemented in codebase: **210**
- The "Implemented in codebase" column is synced to current scalar/text/date/logical builtin registrations, lookup-range builtin registrations, and JS-special formula runtime entries.

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
| ADDRESS | Placeholder (blocked) | Yes | Yes |
| AGGREGATE | Not in ODF 1.4 mandatory | Yes | Yes |
| AMORDEGRC | Not in ODF 1.4 mandatory | Yes | Yes |
| AMORLINC | Implemented | Yes | Yes |
| AND | Not in ODF 1.4 mandatory | Yes | Yes |
| ARABIC | Implemented | Yes | Yes |
| AREAS | Placeholder (blocked) | Yes | Yes |
| ARRAYTOTEXT | Not in ODF 1.4 mandatory | Yes | Yes |
| ASC | Placeholder (blocked) | Yes | No |
| ASIN | Not in ODF 1.4 mandatory | Yes | Yes |
| ASINH | Implemented | Yes | Yes |
| ATAN | Not in ODF 1.4 mandatory | Yes | Yes |
| ATAN2 | Not in ODF 1.4 mandatory | Yes | Yes |
| ATANH | Implemented | Yes | Yes |
| AVEDEV | Placeholder (blocked) | Yes | Yes |
| AVERAGE | Not in ODF 1.4 mandatory | Yes | Yes |
| AVERAGEA | Placeholder (blocked) | Yes | Yes |
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
| BINOM.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| BINOM.DIST.RANGE | Placeholder (blocked) | Yes | Yes |
| BINOM.INV | Not in ODF 1.4 mandatory | Yes | Yes |
| BINOMDIST | Placeholder (blocked) | Yes | Yes |
| BITAND | Placeholder (blocked) | Yes | Yes |
| BITLSHIFT | Placeholder (blocked) | Yes | Yes |
| BITOR | Placeholder (blocked) | Yes | Yes |
| BITRSHIFT | Placeholder (blocked) | Yes | Yes |
| BITXOR | Placeholder (blocked) | Yes | Yes |
| BYCOL | Not in ODF 1.4 mandatory | Yes | Yes |
| BYROW | Not in ODF 1.4 mandatory | Yes | Yes |
| CALL | Not in ODF 1.4 mandatory | Yes | No |
| CEILING | Implemented | Yes | Yes |
| CEILING.MATH | Not in ODF 1.4 mandatory | Yes | Yes |
| CEILING.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| CELL | Not in ODF 1.4 mandatory | Yes | No |
| CHAR | Placeholder (blocked) | Yes | Yes |
| CHIDIST | Not in ODF 1.4 mandatory | Yes | Yes |
| CHIINV | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| CHISQ.DIST.RT | Not in ODF 1.4 mandatory | Yes | Yes |
| CHISQ.INV | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.INV.RT | Not in ODF 1.4 mandatory | Yes | No |
| CHISQ.TEST | Not in ODF 1.4 mandatory | Yes | No |
| CHISQDIST | Missing | No | No |
| CHISQINV | Missing | No | No |
| CHITEST | Not in ODF 1.4 mandatory | Yes | No |
| CHOOSE | Not in ODF 1.4 mandatory | Yes | Yes |
| CHOOSECOLS | Not in ODF 1.4 mandatory | Yes | Yes |
| CHOOSEROWS | Not in ODF 1.4 mandatory | Yes | Yes |
| CLEAN | Placeholder (blocked) | Yes | Yes |
| CODE | Placeholder (blocked) | Yes | Yes |
| COLUMN | Placeholder (blocked) | Yes | No |
| COLUMNS | Not in ODF 1.4 mandatory | Yes | Yes |
| COMBIN | Implemented | Yes | Yes |
| COMBINA | Implemented | Yes | Yes |
| COMPLEX | Placeholder (blocked) | Yes | No |
| CONCAT | Not in ODF 1.4 mandatory | Yes | Yes |
| CONCATENATE | Placeholder (blocked) | Yes | Yes |
| CONFIDENCE | Placeholder (blocked) | Yes | Yes |
| CONFIDENCE.NORM | Not in ODF 1.4 mandatory | Yes | Yes |
| CONFIDENCE.T | Not in ODF 1.4 mandatory | Yes | No |
| CONVERT | Placeholder (blocked) | Yes | No |
| CORREL | Placeholder (blocked) | Yes | Yes |
| COS | Not in ODF 1.4 mandatory | Yes | Yes |
| COSH | Implemented | Yes | Yes |
| COT | Implemented | Yes | Yes |
| COTH | Implemented | Yes | Yes |
| COUNT | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTA | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTBLANK | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTIF | Not in ODF 1.4 mandatory | Yes | Yes |
| COUNTIFS | Implemented | Yes | Yes |
| COUPDAYBS | Placeholder (blocked) | Yes | No |
| COUPDAYS | Placeholder (blocked) | Yes | No |
| COUPDAYSNC | Placeholder (blocked) | Yes | No |
| COUPNCD | Placeholder (blocked) | Yes | No |
| COUPNUM | Placeholder (blocked) | Yes | No |
| COUPPCD | Placeholder (blocked) | Yes | No |
| COVAR | Placeholder (blocked) | Yes | Yes |
| COVARIANCE.P | Not in ODF 1.4 mandatory | Yes | Yes |
| COVARIANCE.S | Not in ODF 1.4 mandatory | Yes | Yes |
| CRITBINOM | Placeholder (blocked) | Yes | Yes |
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
| DATEVALUE | Placeholder (blocked) | Yes | Yes |
| DAVERAGE | Not in ODF 1.4 mandatory | Yes | No |
| DAY | Not in ODF 1.4 mandatory | Yes | Yes |
| DAYS | Placeholder (blocked) | Yes | Yes |
| DAYS360 | Placeholder (blocked) | Yes | Yes |
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
| DELTA | Placeholder (blocked) | Yes | Yes |
| DETECTLANGUAGE | Not in ODF 1.4 mandatory | Yes | No |
| DEVSQ | Placeholder (blocked) | Yes | Yes |
| DGET | Not in ODF 1.4 mandatory | Yes | No |
| DISC | Placeholder (blocked) | Yes | No |
| DMAX | Not in ODF 1.4 mandatory | Yes | No |
| DMIN | Not in ODF 1.4 mandatory | Yes | No |
| DOLLAR | Not in ODF 1.4 mandatory | Yes | Yes |
| DOLLARDE | Placeholder (blocked) | Yes | Yes |
| DOLLARFR | Placeholder (blocked) | Yes | Yes |
| DPRODUCT | Not in ODF 1.4 mandatory | Yes | No |
| DROP | Not in ODF 1.4 mandatory | Yes | Yes |
| DSTDEV | Not in ODF 1.4 mandatory | Yes | No |
| DSTDEVP | Not in ODF 1.4 mandatory | Yes | No |
| DSUM | Not in ODF 1.4 mandatory | Yes | No |
| DURATION | Placeholder (blocked) | Yes | No |
| DVAR | Not in ODF 1.4 mandatory | Yes | No |
| DVARP | Not in ODF 1.4 mandatory | Yes | No |
| EDATE | Implemented | Yes | Yes |
| EFFECT | Placeholder (blocked) | Yes | Yes |
| ENCODEURL | Not in ODF 1.4 mandatory | Yes | Yes |
| EOMONTH | Implemented | Yes | Yes |
| ERF | Placeholder (blocked) | Yes | Yes |
| ERF.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| ERFC | Placeholder (blocked) | Yes | Yes |
| ERFC.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| ERROR.TYPE | Placeholder (blocked) | Yes | Yes |
| EUROCONVERT | Placeholder (blocked) | Yes | No |
| EVEN | Not in ODF 1.4 mandatory | Yes | Yes |
| EXACT | Not in ODF 1.4 mandatory | Yes | Yes |
| EXP | Not in ODF 1.4 mandatory | Yes | Yes |
| EXPAND | Not in ODF 1.4 mandatory | Yes | No |
| EXPON.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| EXPONDIST | Placeholder (blocked) | Yes | Yes |
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
| FINDB | Missing | Yes | Yes |
| FINV | Placeholder (blocked) | Yes | No |
| FISHER | Placeholder (blocked) | Yes | Yes |
| FISHERINV | Placeholder (blocked) | Yes | Yes |
| FIXED | Placeholder (blocked) | Yes | Yes |
| FLOOR | Implemented | Yes | Yes |
| FLOOR.MATH | Not in ODF 1.4 mandatory | Yes | Yes |
| FLOOR.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| FORECAST | Placeholder (blocked) | No | No |
| FORMULA | Missing | No | No |
| FORMULATEXT | Not in ODF 1.4 mandatory | Yes | No |
| FREQUENCY | Placeholder (blocked) | Yes | No |
| FTEST | Placeholder (blocked) | Yes | No |
| FV | Not in ODF 1.4 mandatory | Yes | Yes |
| FVSCHEDULE | Placeholder (blocked) | Yes | No |
| GAMMA | Placeholder (blocked) | Yes | Yes |
| GAMMA.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| GAMMA.INV | Not in ODF 1.4 mandatory | Yes | No |
| GAMMADIST | Placeholder (blocked) | Yes | Yes |
| GAMMAINV | Placeholder (blocked) | Yes | No |
| GAMMALN | Placeholder (blocked) | Yes | Yes |
| GAMMALN.PRECISE | Not in ODF 1.4 mandatory | Yes | Yes |
| GAUSS | Placeholder (blocked) | Yes | Yes |
| GCD | Implemented | Yes | Yes |
| GEOMEAN | Placeholder (blocked) | Yes | Yes |
| GESTEP | Placeholder (blocked) | Yes | Yes |
| GETPIVOTDATA | Placeholder (blocked) | Yes | No |
| GROUPBY | Not in ODF 1.4 mandatory | Yes | No |
| GROWTH | Placeholder (blocked) | Yes | No |
| HARMEAN | Placeholder (blocked) | Yes | Yes |
| HEX2BIN | Placeholder (blocked) | Yes | No |
| HEX2DEC | Placeholder (blocked) | Yes | No |
| HEX2OCT | Placeholder (blocked) | Yes | No |
| HLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| HOUR | Not in ODF 1.4 mandatory | Yes | Yes |
| HSTACK | Not in ODF 1.4 mandatory | Yes | Yes |
| HYPERLINK | Placeholder (blocked) | Yes | No |
| HYPGEOM.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| HYPGEOMDIST | Placeholder (blocked) | Yes | Yes |
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
| IPMT | Placeholder (blocked) | Yes | Yes |
| IRR | Not in ODF 1.4 mandatory | Yes | No |
| ISEVEN | Placeholder (blocked) | Yes | Yes |
| ISFORMULA | Placeholder (blocked) | Yes | Yes |
| ISO.CEILING | Not in ODF 1.4 mandatory | Yes | Yes |
| ISODD | Placeholder (blocked) | No | Yes |
| ISOMITTED | Not in ODF 1.4 mandatory | Yes | No |
| ISOWEEKNUM | Placeholder (blocked) | Yes | Yes |
| ISPMT | Placeholder (blocked) | Yes | Yes |
| ISREF | Placeholder (blocked) | No | Yes |
| JIS | Placeholder (blocked) | No | No |
| KURT | Placeholder (blocked) | Yes | Yes |
| LAMBDA | Not in ODF 1.4 mandatory | Yes | Yes |
| LARGE | Placeholder (blocked) | Yes | Yes |
| LCM | Implemented | Yes | Yes |
| LEFT | Not in ODF 1.4 mandatory | Yes | Yes |
| LEFTB | Missing | Yes | Yes |
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
| LOGINV | Placeholder (blocked) | Yes | Yes |
| LOGNORM.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| LOGNORM.INV | Not in ODF 1.4 mandatory | Yes | Yes |
| LOGNORMDIST | Placeholder (blocked) | Yes | Yes |
| LOOKUP | Placeholder (blocked) | Yes | Yes |
| LOWER | Not in ODF 1.4 mandatory | Yes | Yes |
| MAKEARRAY | Not in ODF 1.4 mandatory | Yes | Yes |
| MAP | Not in ODF 1.4 mandatory | Yes | Yes |
| MATCH | Not in ODF 1.4 mandatory | Yes | Yes |
| MAX | Not in ODF 1.4 mandatory | Yes | Yes |
| MAXA | Placeholder (blocked) | Yes | Yes |
| MAXIFS | Not in ODF 1.4 mandatory | Yes | Yes |
| MDETERM | Implemented | Yes | Yes |
| MDURATION | Placeholder (blocked) | Yes | No |
| MEDIAN | Placeholder (blocked) | Yes | Yes |
| MID | Not in ODF 1.4 mandatory | Yes | Yes |
| MIDB | Missing | Yes | Yes |
| MIN | Not in ODF 1.4 mandatory | Yes | Yes |
| MINA | Placeholder (blocked) | Yes | Yes |
| MINIFS | Not in ODF 1.4 mandatory | Yes | Yes |
| MINUTE | Not in ODF 1.4 mandatory | Yes | Yes |
| MINVERSE | Implemented | Yes | Yes |
| MIRR | Placeholder (blocked) | Yes | No |
| MMULT | Implemented | Yes | Yes |
| MOD | Not in ODF 1.4 mandatory | Yes | Yes |
| MODE | Placeholder (blocked) | Yes | Yes |
| MODE.MULT | Not in ODF 1.4 mandatory | Yes | No |
| MODE.SNGL | Not in ODF 1.4 mandatory | Yes | Yes |
| MONTH | Not in ODF 1.4 mandatory | Yes | Yes |
| MROUND | Implemented | Yes | Yes |
| MULTINOMIAL | Implemented | Yes | Yes |
| MULTIPLE.OPERATIONS | Missing | No | No |
| MUNIT | Implemented | Yes | Yes |
| N | Not in ODF 1.4 mandatory | Yes | Yes |
| NA | Not in ODF 1.4 mandatory | Yes | Yes |
| NEGBINOM.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| NEGBINOMDIST | Placeholder (blocked) | Yes | Yes |
| NETWORKDAYS | Placeholder (blocked) | Yes | Yes |
| NETWORKDAYS.INTL | Not in ODF 1.4 mandatory | Yes | No |
| NOMINAL | Placeholder (blocked) | Yes | Yes |
| NORM.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| NORM.INV | Not in ODF 1.4 mandatory | Yes | Yes |
| NORM.S.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| NORM.S.INV | Not in ODF 1.4 mandatory | Yes | Yes |
| NORMDIST | Placeholder (blocked) | Yes | Yes |
| NORMINV | Placeholder (blocked) | Yes | Yes |
| NORMSDIST | Not in ODF 1.4 mandatory | Yes | Yes |
| NORMSINV | Not in ODF 1.4 mandatory | Yes | Yes |
| NOT | Not in ODF 1.4 mandatory | Yes | Yes |
| NOW | Not in ODF 1.4 mandatory | Yes | Yes |
| NPER | Not in ODF 1.4 mandatory | Yes | Yes |
| NPV | Not in ODF 1.4 mandatory | Yes | Yes |
| NUMBERVALUE | Placeholder (blocked) | Yes | No |
| OCT2BIN | Placeholder (blocked) | Yes | No |
| OCT2DEC | Placeholder (blocked) | Yes | No |
| OCT2HEX | Placeholder (blocked) | Yes | No |
| ODD | Not in ODF 1.4 mandatory | Yes | Yes |
| ODDFPRICE | Placeholder (blocked) | Yes | No |
| ODDFYIELD | Placeholder (blocked) | Yes | No |
| ODDLPRICE | Placeholder (blocked) | Yes | No |
| ODDLYIELD | Placeholder (blocked) | Yes | No |
| OFFSET | Placeholder (blocked) | Yes | Yes |
| OR | Not in ODF 1.4 mandatory | Yes | Yes |
| PDURATION | Placeholder (blocked) | Yes | Yes |
| PEARSON | Placeholder (blocked) | Yes | Yes |
| PERCENTILE | Placeholder (blocked) | Yes | No |
| PERCENTILE.EXC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTILE.INC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTOF | Not in ODF 1.4 mandatory | Yes | Yes |
| PERCENTRANK | Placeholder (blocked) | Yes | No |
| PERCENTRANK.EXC | Not in ODF 1.4 mandatory | Yes | No |
| PERCENTRANK.INC | Not in ODF 1.4 mandatory | Yes | No |
| PERMUT | Placeholder (blocked) | Yes | Yes |
| PERMUTATIONA | Placeholder (blocked) | Yes | Yes |
| PHI | Placeholder (blocked) | Yes | Yes |
| PHONETIC | Not in ODF 1.4 mandatory | Yes | No |
| PI | Not in ODF 1.4 mandatory | Yes | Yes |
| PIVOTBY | Not in ODF 1.4 mandatory | Yes | No |
| PMT | Not in ODF 1.4 mandatory | Yes | Yes |
| POISSON | Placeholder (blocked) | Yes | Yes |
| POISSON.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| POWER | Not in ODF 1.4 mandatory | Yes | Yes |
| PPMT | Placeholder (blocked) | Yes | Yes |
| PRICE | Placeholder (blocked) | Yes | No |
| PRICEDISC | Placeholder (blocked) | Yes | No |
| PRICEMAT | Placeholder (blocked) | Yes | No |
| PROB | Placeholder (blocked) | Yes | No |
| PRODUCT | Not in ODF 1.4 mandatory | Yes | Yes |
| PROPER | Not in ODF 1.4 mandatory | Yes | Yes |
| PV | Not in ODF 1.4 mandatory | Yes | Yes |
| QUARTILE | Placeholder (blocked) | Yes | No |
| QUARTILE.EXC | Not in ODF 1.4 mandatory | Yes | No |
| QUARTILE.INC | Not in ODF 1.4 mandatory | Yes | No |
| QUOTIENT | Implemented | Yes | Yes |
| RADIANS | Not in ODF 1.4 mandatory | Yes | Yes |
| RAND | Implemented | Yes | Yes |
| RANDARRAY | Not in ODF 1.4 mandatory | Yes | Yes |
| RANDBETWEEN | Implemented | Yes | Yes |
| RANK | Placeholder (blocked) | Yes | Yes |
| RANK.AVG | Not in ODF 1.4 mandatory | Yes | No |
| RANK.EQ | Not in ODF 1.4 mandatory | Yes | Yes |
| RATE | Not in ODF 1.4 mandatory | Yes | No |
| RECEIVED | Placeholder (blocked) | Yes | No |
| REDUCE | Not in ODF 1.4 mandatory | Yes | Yes |
| REGEXEXTRACT | Not in ODF 1.4 mandatory | Yes | No |
| REGEXREPLACE | Not in ODF 1.4 mandatory | Yes | No |
| REGEXTEST | Not in ODF 1.4 mandatory | Yes | No |
| REGISTER.ID | Not in ODF 1.4 mandatory | Yes | No |
| REPLACE | Not in ODF 1.4 mandatory | Yes | Yes |
| REPLACEB | Missing | Yes | No |
| REPT | Not in ODF 1.4 mandatory | Yes | Yes |
| RIGHT | Not in ODF 1.4 mandatory | Yes | Yes |
| RIGHTB | Missing | Yes | Yes |
| ROMAN | Implemented | Yes | Yes |
| ROUND | Not in ODF 1.4 mandatory | Yes | Yes |
| ROUNDDOWN | Implemented | Yes | Yes |
| ROUNDUP | Implemented | Yes | Yes |
| ROW | Placeholder (blocked) | Yes | No |
| ROWS | Not in ODF 1.4 mandatory | Yes | Yes |
| RRI | Placeholder (blocked) | Yes | Yes |
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
| SKEW | Placeholder (blocked) | Yes | Yes |
| SKEW.P | Not in ODF 1.4 mandatory | Yes | Yes |
| SKEWP | Missing | No | No |
| SLN | Not in ODF 1.4 mandatory | Yes | No |
| SLOPE | Placeholder (blocked) | Yes | No |
| SMALL | Placeholder (blocked) | Yes | Yes |
| SORT | Not in ODF 1.4 mandatory | Yes | Yes |
| SORTBY | Not in ODF 1.4 mandatory | Yes | Yes |
| SQRT | Not in ODF 1.4 mandatory | Yes | Yes |
| SQRTPI | Implemented | Yes | Yes |
| STANDARDIZE | Placeholder (blocked) | Yes | Yes |
| STDEV | Not in ODF 1.4 mandatory | Yes | Yes |
| STDEV.P | Not in ODF 1.4 mandatory | Yes | Yes |
| STDEV.S | Not in ODF 1.4 mandatory | Yes | Yes |
| STDEVA | Placeholder (blocked) | Yes | Yes |
| STDEVP | Not in ODF 1.4 mandatory | Yes | Yes |
| STDEVPA | Placeholder (blocked) | Yes | Yes |
| STEYX | Placeholder (blocked) | Yes | No |
| STOCKHISTORY | Not in ODF 1.4 mandatory | Yes | No |
| SUBSTITUTE | Not in ODF 1.4 mandatory | Yes | Yes |
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
| T | Not in ODF 1.4 mandatory | Yes | Yes |
| T.DIST | Not in ODF 1.4 mandatory | Yes | No |
| T.DIST.2T | Not in ODF 1.4 mandatory | Yes | No |
| T.DIST.RT | Not in ODF 1.4 mandatory | Yes | No |
| T.INV | Not in ODF 1.4 mandatory | Yes | No |
| T.INV.2T | Not in ODF 1.4 mandatory | Yes | No |
| T.TEST | Not in ODF 1.4 mandatory | Yes | No |
| TAKE | Not in ODF 1.4 mandatory | Yes | Yes |
| TAN | Not in ODF 1.4 mandatory | Yes | Yes |
| TANH | Implemented | Yes | Yes |
| TBILLEQ | Placeholder (blocked) | Yes | No |
| TBILLPRICE | Placeholder (blocked) | Yes | No |
| TBILLYIELD | Placeholder (blocked) | Yes | No |
| TDIST | Not in ODF 1.4 mandatory | Yes | No |
| TEXT | Placeholder (blocked) | Yes | No |
| TEXTAFTER | Not in ODF 1.4 mandatory | Yes | Yes |
| TEXTBEFORE | Not in ODF 1.4 mandatory | Yes | Yes |
| TEXTJOIN | Not in ODF 1.4 mandatory | Yes | Yes |
| TEXTSPLIT | Not in ODF 1.4 mandatory | Yes | No |
| TIME | Not in ODF 1.4 mandatory | Yes | Yes |
| TIMEVALUE | Placeholder (blocked) | Yes | Yes |
| TINV | Placeholder (blocked) | Yes | No |
| TOCOL | Not in ODF 1.4 mandatory | Yes | Yes |
| TODAY | Not in ODF 1.4 mandatory | Yes | Yes |
| TOROW | Not in ODF 1.4 mandatory | Yes | Yes |
| TRANSLATE | Not in ODF 1.4 mandatory | Yes | No |
| TRANSPOSE | Placeholder (blocked) | Yes | Yes |
| TREND | Placeholder (blocked) | Yes | No |
| TRIM | Not in ODF 1.4 mandatory | Yes | Yes |
| TRIMMEAN | Placeholder (blocked) | Yes | No |
| TRIMRANGE | Not in ODF 1.4 mandatory | Yes | No |
| TRUE | Not in ODF 1.4 mandatory | Yes | Yes |
| TRUNC | Not in ODF 1.4 mandatory | Yes | Yes |
| TTEST | Placeholder (blocked) | Yes | No |
| TYPE | Placeholder (blocked) | Yes | Yes |
| UNICHAR | Placeholder (blocked) | Yes | Yes |
| UNICODE | Placeholder (blocked) | Yes | Yes |
| UNIQUE | Not in ODF 1.4 mandatory | Yes | Yes |
| UPPER | Not in ODF 1.4 mandatory | Yes | Yes |
| USE.THE.COUNTIF | Not in ODF 1.4 mandatory | Yes | No |
| VALUE | Not in ODF 1.4 mandatory | Yes | Yes |
| VALUETOTEXT | Not in ODF 1.4 mandatory | Yes | No |
| VAR | Not in ODF 1.4 mandatory | Yes | Yes |
| VAR.P | Not in ODF 1.4 mandatory | Yes | Yes |
| VAR.S | Not in ODF 1.4 mandatory | Yes | Yes |
| VARA | Placeholder (blocked) | Yes | Yes |
| VARP | Not in ODF 1.4 mandatory | Yes | Yes |
| VARPA | Placeholder (blocked) | Yes | Yes |
| VDB | Placeholder (blocked) | Yes | No |
| VLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| VSTACK | Not in ODF 1.4 mandatory | Yes | Yes |
| WEBSERVICE | Not in ODF 1.4 mandatory | Yes | No |
| WEEKDAY | Not in ODF 1.4 mandatory | Yes | Yes |
| WEEKNUM | Placeholder (blocked) | Yes | Yes |
| WEIBULL | Placeholder (blocked) | Yes | Yes |
| WEIBULL.DIST | Not in ODF 1.4 mandatory | Yes | Yes |
| WORKDAY | Placeholder (blocked) | Yes | Yes |
| WORKDAY.INTL | Not in ODF 1.4 mandatory | Yes | No |
| WRAPCOLS | Not in ODF 1.4 mandatory | Yes | Yes |
| WRAPROWS | Not in ODF 1.4 mandatory | Yes | Yes |
| XIRR | Placeholder (blocked) | Yes | No |
| XLOOKUP | Not in ODF 1.4 mandatory | Yes | Yes |
| XMATCH | Not in ODF 1.4 mandatory | Yes | Yes |
| XNPV | Placeholder (blocked) | Yes | No |
| XOR | Implemented | Yes | Yes |
| YEAR | Not in ODF 1.4 mandatory | Yes | Yes |
| YEARFRAC | Placeholder (blocked) | Yes | Yes |
| YIELD | Placeholder (blocked) | Yes | No |
| YIELDDISC | Placeholder (blocked) | Yes | No |
| YIELDMAT | Placeholder (blocked) | Yes | No |
| Z.TEST | Not in ODF 1.4 mandatory | Yes | No |
| ZTEST | Placeholder (blocked) | Yes | No |
