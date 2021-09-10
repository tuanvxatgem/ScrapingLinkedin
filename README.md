# ScrapingLinkedin

## Perrequisites
Install last version Google Chrome and python
## Installing:
pip install -r requirements.txt

Run `LinkedinScraping\configurator.py` following the prompted instructions. Or change in file config.ini
## To run
python do_scraping2.py
If you need run with dont open a real Chrome window:
### Headless execution
In this mode the script will do scraping without opening a real Chrome window.

**Pros:** The scraping process is distributed into many threads to speed up to 4 times the performance. Moreover, in this way you can keep on doing your regular business on your computer as you don't have to keep the focus on any specific window.

**Cons:** If you scrap many profiles (more than hundreds) and/or in unusual times (in the night) LinkedIn may prompt a Captcha to check that you are not a human. If this happens, there is no way for you to fill in the Captcha. The script will detect this particular situation and terminate the scraping with an alert: you will have hence to run the script in normal mode, do the captcha, and then you can proceed in scraping the profiles that were left.

python do_scraping2.py HEADLESS

## Disclaimer

The repository is intended to be used as a reference to learn more on Python and to perform scraping for personal usage. Every country has different and special regulations on usage of personal information, so I strongly recommend you to check your national legislation before using / sharing / selling / elaborating the scraped information. I decline any responsibility on the usage of scraped information. Cloning this repository and executing the included scripts you declare and confirm the responsibility of the scraped data usage is totally up on you.
