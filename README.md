This program is designed to capture financial statements of those corporations which are listed in HKEX and transform them into XLSX. It is well known that a plenty of financial data softwares such as Wind or Bloomberg have collected countless financial data and charts including significant financial stataments such as the BALANCE SHEET,INCOME STATEMENT and so on in the format of XLS, which implies that it seems meaningless to develop another program which has quite similar functions. However, noticing that those financial statements released in HKEX are always attached with NUMEROUS notes in which a lot of charts are displayed but not collected in those famous existing databases while those notes do matter when estimating the situtation and prospect of corporations, it is quite necessary to transform those financial statements displayed in the notes into STRUCTURAL DATA such as tables in the format of XLSX to make them easier to be processed by analysts.

NOTES:

0. The version having been launched is named to be BETA 2.0 suggesting that the program may collapse or be not able to work in high efficieny under some extreme conditions, in spite of the fact that the program has performed quite well in most tests.

1. The program may not be able to produce perfect XLSX sometimes, thus artificial interventions and adjustments are necessary when the program can't perform perfectly.

2. Some packages are necessary to assure that the program is able to run well, including:

(1) BeautifulSoup 4(bs 4)

(2) re

(3) openpyxl

(4) requests

(5) selenium

Those packages can be easily installed through PIP.
Most importantly, the Internet browser launched by google i.e. CHROME and one of its plugins i.e. CHROMEDRIVER are necessary. You have to check whether the version of chromedriver is the latest or not. The version of chromedriver may be too old to drive chrome and the program can be ineffective consequently.

3. Please use Python 3.X.

4. The Internet has to be available.

5. The program will be updated if the developer finds a more effective algorithm.# financial-statements-TO-xlsx
capture financial statements of those corporations which are listed in HKEX and transform them into XLSX.
