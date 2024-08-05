import sys
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import datetime
import random
import string
import difflib
import re
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)

# Initialize the speaker
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Voice = speaker.GetVoices().Item(1)

def takeCmnd():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-US")
            query = replace_special_characters(query)
            print(f"User Said: {query}")
            return query
        except sr.UnknownValueError:
            say("Sorry, I didn't catch that. Please repeat.")
            return None
        except sr.RequestError:
            say("Network error. Please check your internet connection.")
            return None

special_character_mapping = {
    "dot": ".",
    "comma": ",",
    "exclamation": "!",
    "number": "#",
    "dash": "-",
    " ": ""
}

def replace_special_characters(text):
    words = text.split()
    replaced_words = []
    for word in words:
        replaced_words.append(special_character_mapping.get(word, word))
    return " ".join(replaced_words)

def say(text):
    print("Speak.")
    speaker.Speak(text)

def find_file(base_filename):
    for root, dirs, files in os.walk(os.path.expanduser("~")):
        for file in files:
            if file.lower().startswith(base_filename.lower()):
                file_path = os.path.join(root, file)
                os.startfile(file_path)
                print(f"Opening {file} at {file_path}")
                return True
    print(f"File {base_filename} not found.")
    return False

def create_search_space(directory):
    search_space = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            search_space.append(file)
    return search_space

def levenshtein_distance(s1, s2):
    return sum(1 for s in difflib.ndiff(s1, s2) if s[0] != ' ')

def genetic_algorithm_search(target_file, search_space, population_size=50, generations=100):
    def generate_random_string(length):
        characters = string.ascii_letters + string.digits + "._-"
        return ''.join(random.choice(characters) for _ in length)

    def mutate_string(s, mutation_rate):
        mutated = [c if random.random() > mutation_rate else random.choice(string.ascii_letters + string.digits + "._-") for c in s]
        return ''.join(mutated)

    search_directory = "D:\\"
    search_paths = [os.path.join(search_directory, filename) for filename in search_space]

    population = [generate_random_string(len(target_file)) for _ in population_size]

    for generation in range(generations):
        population = sorted(population, key=lambda x: levenshtein_distance(target_file, x))
        best_individual = population[0]

        if best_individual == target_file:
            for path in search_paths:
                if target_file.lower() in path.lower():
                    return path

        new_population = [best_individual]

        for _ in range(population_size - 1):
            parent1 = random.choice(population[:population_size // 2])
            parent2 = random.choice(population[:population_size // 2])
            child1, child2 = crossover(parent1, parent2)
            new_population.append(mutate_string(child1, mutation_rate=0.2))
            new_population.append(mutate_string(child2, mutation_rate=0.2))

        population = new_population

    return None

def crossover(parent1, parent2):
    crossover_point = random.randint(1, len(parent1) - 1)
    child1 = parent1[:crossover_point] + parent2[crossover_point:]
    child2 = parent2[:crossover_point] + parent1[crossover_point:]
    return child1, child2

def open_website(website_name):
    url = f"https://www.{website_name}.com"
    webbrowser.open(url)
    print(f"Opening {website_name}: {url}")

# def confirm_command(command):
#     say(f"Did you say {command}?")
#     confirmation = takeCmnd()
#     if confirmation and re.search(r'\byes\b', confirmation, re.IGNORECASE):
#         return True
#     return False

commands = {
    "open": {
        "websites": {
            "youtube": "https://youtube.com",
            "google": "https://google.com",
            "wikipedia": "https://wikipedia.com",
            "chat gpt": "https://chat.openai.com"
        },
        "apps": {
            "visual studio code": r"C:\Users\Rumman\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Visual Studio Code",
            "zoom": r"C:\Users\Rumman\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Zoom",
            "music": r"C:\Users\HP\Downloads\tvari-hawaii-vacation-159069.mp3"
        }
    }
}

if __name__ == '__main__':
    print('Welcome.')
    say("Hello, I am Emma.")
    while True:
        query = takeCmnd()

        if query:
            for cmd_type, cmd_details in commands.items():
                for category, items in cmd_details.items():
                    for item_name, item_path in items.items():
                        if re.search(fr'\b{cmd_type} {item_name}\b', query, re.IGNORECASE):
                            if cmd_type == "open" and category == "websites":
                                webbrowser.open(item_path)
                            elif cmd_type == "open" and category == "apps":
                                os.startfile(item_path)

            if re.search(r'\bthe time\b', query, re.IGNORECASE):
                Time = datetime.datetime.now().strftime("%H:%M:%S")
                say(f"The time is {Time}")

            if re.search(r'\bopen file\b', query, re.IGNORECASE):
                query = query.replace("open file", "").strip()
                say("Sure, what's the name of the file you want to open?")
                file_name = takeCmnd()

                if file_name:
                    if "genetic algo" in query.lower():
                        search_space = create_search_space("D:\\")
                        found_path = genetic_algorithm_search(file_name, search_space)
                        if found_path:
                            print("Using gen algo, found path:", found_path)
                            try:
                                os.startfile(found_path)
                                say(f"Using gen algo, file opened: {found_path}")
                            except FileNotFoundError:
                                say(f"Sorry, using gen algo, the file '{file_name}' could not be found.")
                        else:
                            say("Using gen algo, file not found.")
                    else:
                        found_path = find_file(file_name)
                        if found_path:
                            print("Found path:", found_path)
                            try:
                                os.startfile(found_path)
                                say(f"File opened: {found_path}")
                            except FileNotFoundError:
                                say(f"Sorry, the file '{file_name}' could not be found.")
                        else:
                            say("File not found.")

            elif re.search(r'\bclose file\b', query, re.IGNORECASE):
                base_filename = query.replace("close file", "").strip()
                print(f"Closing {base_filename}.")
                say(f"Closing {base_filename}.")
                os.system(f"taskkill /im {base_filename}.exe /f")

            elif re.search(r'\bopen website\b', query, re.IGNORECASE):
                website_name = query.replace("open website", "").strip()
                say(f"Opening {website_name}.")
                open_website(website_name)

            elif re.search(r'\bstop\b', query, re.IGNORECASE):
                say("Good Bye.")
                sys.exit()

