# InsiderMiner

InsiderMiner was a stock forecasting system built between 2010 and 2014 by William B. Claster, co-founder, CTO and chief scientist. Each trading morning it predicted whether a stock would close up, down or unchanged, using the volume of social-media discussion about the company together with recent price behavior. The predictions were published through Android and iPhone apps.

This repository preserves the original source code. It is published as a historical record and is not maintained.

## How the system worked

1. **Tweet collection.** A Perl script collected tweets mentioning each tracked company and saved them as daily CSV files.
2. **Feature construction (Excel VBA).** A set of macros imported the daily files, estimated tweet counts, normalized each company's count by the total tweet volume for the day, and shifted tweet times so that only tweets posted before the market opened counted toward that day. Weekend activity was folded into the next trading day. The macros then joined the counts with daily prices and built the model inputs: close minus open, the day's range, today's open minus yesterday's close, normalized volume, the ratio of tweet count to volume, interaction terms, and up to 20 days of lagged tweet counts and price fields. The target label was U, D or N.
3. **Prediction (Java and Weka).** A Java program read the training table from a MySQL database, trained a Weka J48 decision tree, classified the current day, and wrote the prediction and the latest price back to the database at about 8:30 a.m. New York time.
4. **Delivery (mobile app).** A PhoneGap app for Android and iPhone read the prediction table and showed an up, down or no-change forecast for each stock. Free and paid versions differed in the number of stocks and days per week a user could view.

Stocks tracked at various times included Sony (Tokyo, 6758), Coca-Cola, Disney, Groupon, Microsoft, Verizon, Walmart and Yahoo.

## Contents

| Path | Date | Description |
|---|---|---|
| `excel-macros/AllMacsForInsiderJan142013.bas` | January 2013 | All 34 VBA macros, exported from Excel. Covers import, normalization, time alignment, price download, weekend handling, feature construction and lags. |
| `java/wekaJ48Sony.java` | April 2012 | Daily train, predict and write-back program for Sony using Weka J48 and MySQL. |
| `r/weekend_aggregation.R` | December 2014 | R translation of the time-alignment and weekend-aggregation step. |
| `python/Tpym.py` | May 2016 | A later Python collector that streams tweets, scores sentiment with TextBlob, and stores them in Elasticsearch. |

### Main macros, in pipeline order

| Macro | Purpose |
|---|---|
| `CSV_Import_Loop`, `CSV_Import_TweetsLoop` | Import daily tweet files |
| `generateDatesAndMbFormbPerFileFromPerl`, `fromMbToNumTweets` | Estimate daily tweet counts from file sizes |
| `normalizeCountsWithCountAll` | Normalize by total daily tweet volume |
| `pushBackTimeForAllTweetsNYTime`, `AddTimeandChangeToDateOnlyForJapan` | Align tweet times with the New York or Tokyo trading day |
| `GetData` | Download daily prices for a symbol and date range |
| `ExtendedPriceCountMatchSvet`, `fromSvetToCleanUP2` | Match tweet days with trading days and fold in non-trading days |
| `MakeFieldsNoLagsNYSE` | Build price and count features and the U/D/N label |
| `MakeFirstSetOfLagsNYSE`, `CleanUPMakeSureOfTimeAlignmentNYSE` | Build lagged fields and check time alignment |

## Notes

- Database addresses, user names, passwords and API keys have been removed. The code now reads them from environment variables, and it will not run as published without a database matching the original tables.
- The code is reproduced as written, including working comments. It depends on libraries of its time (Weka 3.7, the Google Finance Portfolio Data API, Python 2, early versions of Tweepy and Elasticsearch), several of which are no longer available.
