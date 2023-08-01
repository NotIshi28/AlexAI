import win32com.client
import speech_recognition as sr
import webbrowser as wb

speaker = win32com.client.Dispatch("Sapi.SpVoice")

def say(i):
    speaker.Speak(i)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occurred. Sorry from Alex"


if __name__ == '__main__':
    say("Hello I am Alex A.I. . What is your name ?")
    print("Listening...")
    name = takeCommand()
    say(f"Hi {name}. what can i do for you")
    print(f"Hi {name}")
    while True:
        print("Listening...")
        query = takeCommand()
        sites = [["youtube", "https://youtube.com"],
                 ["facebook", "https://facebook.com"],
                 ["whatsapp", "https://web.whatsapp.com/"],
                 ["google", "https://google.com"],
                 ["chat GPT", "https://chat.openai.com/"],
                 ["twitter", "https://twitter.com/"],
                 ["instagram", "https://www.instagram.com/"],
                 ['discord', 'https://discord.com/'],
                 ['spotify', 'https://open.spotify.com/'],
                 ['makemytrip', 'https://www.makemytrip.com/'],
                 ['flipkart', 'https://www.flipkart.com/'],
                 ['amazon', 'https://www.amazon.in/'],
                 ['netflix', 'https://www.netflix.com/'],
                 ]
        for site in sites:
            if f"open {site[0]}" in query.lower():
                say(f"Opening {site[0]} sir ...")
                wb.open(site[1])
                query = ""
            elif "I am planning a solo trip which destinations are known to be safe and welcoming for solo travellers".lower() in query.lower():
                query = ""
                say("The safest destinations are in the mountains, the deserts, and the ocean. The safest destinations for solo travelers are in the mountains, the deserts, and the ocean.")
            elif "What are some of the beaten path destinations that are worth exploring".lower() in query.lower():
                query = ""
                say("Exploring lesser-known places can be thrilling! Consider visiting Slovenia's Lake Bled, Cambodia's Kampot, or Portugal's Azores Islands for unique experiences.")
            elif "I am interested in adventure travel".lower() in query.lower():
                query = ""
                say("Adventure awaits! You can try hiking in Peru's Machu Picchu, Nepal's Everest Base Camp, or New Zealand's Tongariro Alpine Crossing.")
            elif "I have a layover in Singapore for a few hours. What are some quick things I can do or see".lower() in query.lower():
                query = ""
                say("Singapore is perfect for short layovers! You can visit Gardens by the Bay, Marina Bay Sands, or take a quick ride on the Singapore Flyer.")
            elif "I want to browse hotels".lower() in query.lower():
                query = ""
                say("You can refer to this website for hotels")
                wb.open("https://www.booking.com/")
            elif "I want to browse flights".lower() in query.lower():
                query = ""
                say("You can refer to this website for flights")
                wb.open("https://www.expedia.com/")
            elif "I want to browse restaurants".lower() in query.lower():
                query = ""
                say("You can refer to this website for restaurants")
                wb.open("https://www.tripadvisor.com/")
            elif "I want to browse movies".lower() in query.lower():
                query = ""
                say("You can refer to this website for movies")
                wb.open("https://www.imdb.com/")
            elif "I want to browse music".lower() in query.lower():
                query = ""
                say("You can refer to this website for music")
                wb.open("https://www.spotify.com/")
                say("I hope you enjoy listening to music.")
            elif  "I want to browse books".lower() in query.lower():
                query = ""
                say("You can refer to this website for books")
                wb.open("https://www.goodreads.com/")
            elif "what are the latest travel destinations".lower() in query.lower():
                query = ""
                say("You can refer to this website for travel destinations")
                wb.open("https://www.lonelyplanet.com/best-in-travel")
            elif "what are the best domestic travel destinations".lower() in query.lower():
                query = ""
                wb.open("https://timesofindia.indiatimes.com/travel/destinations")
            elif "what are the best adventure destinations".lower() in query.lower():
                say("You can refer to this website for adventure destinations")
                query = ""
                wb.open("https://www.cntraveler.com/gallery/the-best-countries-for-adventure-travel")
            elif "what are the most peaceful destinations".lower() in query.lower():
                query = ""
                say("You can refer to this website for peaceful destinations")
                wb.open("https://www.makemytrip.com/tripideas/relaxation-destinations")


            elif "Alex Quit".lower() in query.lower():
                exit()
