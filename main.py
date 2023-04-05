# import necessary libraries
import requests  # for making HTTP requests
import json  # for handling JSON data
import win32com.client as wincl  # for text-to-speech functionality

# continuously prompt for the name of a city and retrieve its weather information until user inputs "q"
while True:
    city = input("Enter the name of the city: \n")  # prompt the user to enter a city name
    if city == "q":  # if user inputs "q", break out of the loop
        # use wincl library to say "thanks for using" when the user quits
        speaker = wincl.Dispatch("SAPI.SpVoice")
        speaker.Speak("Thanks for using!")
        break

    # use the weatherapi API to retrieve current weather information for the specified city
    url = f"http://api.weatherapi.com/v1/current.json?key=c3c16bf5ddce4ad4af8144449230404&q={city}"
    r = requests.get(url)  # make an HTTP GET request to the API endpoint
    wdic = json.loads(r.text)  # parse the response data into a dictionary

    # Retrieve the last_updated value as a string
    last_updated = wdic["current"]["last_updated"]

    # Retrieve the temperature value as a float
    temperature_celsius = wdic['current']['temp_c']

    # Create a text string to be spoken and printed by concatenating the retrieved values
    text_to_speak = f"The current weather in {city} is {temperature_celsius} degrees Celsius. The last update was at {last_updated}."
    print(text_to_speak)  # print the text output to the console

    # Use the wincl library to speak out the text string
    speaker = wincl.Dispatch("SAPI.SpVoice")
    speaker.Speak(text_to_speak)  # use text-to-speech to output the text string