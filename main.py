import tweepy
import sys
import win32com.client as wincl


auth = tweepy.OAuthHandler("API_KEY", "API_ACCESS_KEY")
auth.set_access_token("ACCESS_TOKEN", "ACCESS_TOKEN_SECRET")

api = tweepy.API(auth)

count = len(sys.argv)
global exists
exists = False

speak = wincl.Dispatch("SAPI.SpVoice")

def cleanTweet(text):
	stopwords = ('http','rt','@', '#')
	textwords = text.split()

	resultwords  = [word for word in textwords if not word.lower().startswith(stopwords)]
	result = ' '.join(resultwords)
	return result


public_tweets = api.home_timeline()
if (count == 1):
	for tweet in public_tweets:
		if (tweet.truncated == "true"):
			text = tweet.extended_tweet.encode("utf8")
			text = cleanTweet(text)
			exists = True
			print(text)
			print("------------------------------------------------")
			speak.Speak(text)
		else:
			text = tweet.text.encode("utf8")
			text = cleanTweet(text)
			exists = True
			print(text)
			print("------------------------------------------------")
			speak.Speak(text)
	if (not exists):
		print("No results were found")
elif count == 2:
	for tweet in tweepy.Cursor(api.user_timeline, screen_name=sys.argv[1], tweet_mode="extended").items():
		text = tweet.full_text.encode("utf8")
		text = cleanTweet(text)
		exists = True
		print(text)
		print("------------------------------------------------")
		speak.Speak(text)
		
	if (not exists):
		print("No results were found")