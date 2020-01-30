# bible-verse-auto-fill

#### 1. Auto-Fill Bible Verses using Crossway ESV API : populate_bible_verses_api.py

1. Populate the BOOK, CHAPTER, VERSE, REFERENCE fields in the excel sheet and let the code auto-fill the VERSE_TEXT field using "'https://api.esv.org/v3/passage/text/'"

2. **Libraries used are** :
	1. <ins>For reading and writing XLS workbooks - </ins>pandas
	2. <ins>Other libraries :</ins>
		1. requests
		2. configparser

#### 2. Auto-Fill Bible Verses by Web Scrapping https://www.biblegateway.com : populate_bible_verses.py

1. Populate the BOOK, CHAPTER, VERSE, REFERENCE fields in the excel sheet and let the code auto-fill the VERSE_TEXT field for any Bible versions using Web Scrapping.

2. **Libraries used are** :
	1. <ins>For reading and writing XLS workbooks : </ins>
		1. xlrd
		2. xlwt
	2. <ins>For Web Scrapping : </ins>
		1. BeautifulSoup Version 4
		2. requests