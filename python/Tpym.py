# Twitter streaming sentiment collector, 2016. Credentials removed for publication.
import json
import tweepy
from tweepy.streaming import StreamListener
from tweepy import OAuthHandler
from tweepy import Stream
from textblob import TextBlob
from elasticsearch import Elasticsearch
from config import *
import time
from time import sleep
import urllib3
import os
import sys
import datetime
logf = open("download.log", "w")
GEOBOX_WORLD = [-180,-90,180,90]
urllib3.disable_warnings()
es = Elasticsearch()
access_token = os.environ.get("TWITTER_ACCESS_TOKEN")
access_token_secret = os.environ.get("TWITTER_ACCESS_TOKEN_SECRET")
consumer_key = os.environ.get("TWITTER_CONSUMER_KEY")
consumer_secret = os.environ.get("TWITTER_CONSUMER_SECRET")
class TweetStreamListener(StreamListener):
    def on_data(self, data):
        dict_data = json.loads(data)
        #time.sleep(0.01)
        try: 
            dict_data["text"]
            tweet = TextBlob(dict_data["text"])
            print tweet.sentiment.polarity
            if tweet.sentiment.polarity < 0:
                sentiment = "negative"
            elif tweet.sentiment.polarity == 0:
                sentiment = "neutral"
            else:
                sentiment = "positive"
            print sentiment
            es.index(index="sentiment",
                     doc_type="test-type",
                     body={"author": dict_data["user"]["screen_name"],
                           "favoriteCount": dict_data["favorite_count"],
                           "tweetID": dict_data["id"],
                           "lang": dict_data["lang"],
                           "tweetID": dict_data["id"],                           
                           "date": dict_data["created_at"],
                           "message": dict_data["text"],
                           "polarity": tweet.sentiment.polarity,
                           "subjectivity": tweet.sentiment.subjectivity,
                           "sentiment": sentiment})
        except KeyError:
                pass            
        return True
    def on_error(self, status):
        print status
        print time.now()
        ts = time.time()
        sttime = datetime.datetime.fromtimestamp(ts).strftime('%Y%m%d_%H:%M:%S - ') 
        error = "ERROR! No file 'bla' found!"
        log = 'log.txt'
        logfile = file(log, "w")
        logfile.write(sttime + error + '\n')
        logfile.close()
        sys.exit(0)        
        
if __name__ == '__main__':
    listener = TweetStreamListener()
    auth = OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    stream = Stream(auth, listener)
    #stream.filter(track=['earthquake', 'Trump','Beppu','japancupid'])
    stream.filter(locations=GEOBOX_WORLD)














