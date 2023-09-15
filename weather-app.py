import requests
import json
import win32com.client

def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def get_weather_info(city):
    url = f"https://api.weatherapi.com/v1/current.json?key=38c8d7112b0c45cabae120013231808&q={city}"
    response = requests.get(url)
    data = json.loads(response.text)
    return data

def main():
    city = input("Enter the name of the city of which you want the weather: ")
    weather_data = get_weather_info(city)

    print("available information:\nname, region, country, lat, lon, tz_id, localtime, temp_c, temp_f, is_day, wind_mph, wind_kph, wind_degree, wind_dir, pressure_mb, pressure_in, humidity, cloud, feelslike_c, feelslike_f, vis_km, vis_miles, uv, quit\n")
    speak("available information: name, region, country, latitude, longitude, time-zone, localtime, temperature in celcius, temperature in fahrenheit, day or night, wind speed in meter per hour, wind speed in kilometer per hour, wind_degree, wind's direction, pressure in mb, pressure in inch, humidity, cloud, feelslike in celcius, feelslike in farenheight, visibility in km, visibility in miles, uv radiation, quit")

    while True:
            x = input("What do you want to know about the city ?:")
            if x == "quit":
                print("Thank You!")
                speak("Thank You!")
                break
            elif x in weather_data['location']:
                w = weather_data['location'][x]
                print(f"{x} of {city} is {w}")
                speak(f"{x} of {city} is {w}")
            elif x in weather_data['current']:
                w = weather_data['current'][x]
                if x == 'is_day':
                    print("It's day" if w == 1 else "It's night")
                    speak("It's day" if w == 1 else "It's night")
                else:
                    print(f"{x.replace('_', ' ')} in {city} is {w}")
                    speak(f"{x.replace('_', ' ')} in {city} is {w}")
            else:
                print("Invalid input. Please try again.")

if __name__ == "__main__":
    main()