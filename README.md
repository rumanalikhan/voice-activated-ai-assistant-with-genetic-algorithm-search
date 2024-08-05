## Voice-Activated Assistant
A voice-activated assistant built with Python, capable of performing various tasks such as opening websites, applications, files, and providing the current time. This assistant can recognize special characters in voice commands and use genetic algorithms for efficient file searching.

## Features
**Voice Command Recognition:** Utilizes Google Speech Recognition for interpreting voice commands.
**Website and Application Interaction:** Opens specified websites and applications based on voice commands.
**Time Reporting:** Provides the current time upon request.
**File Operations:** Searches and opens files using standard and genetic algorithm-based methods for enhanced efficiency.
**Special Character Recognition:** Replaces spoken special characters like dot, comma, and dash for improved command accuracy.

## Genetic Algorithm File Search
The assistant employs a genetic algorithm to enhance file search efficiency when regular methods are insufficient. The genetic algorithm mimics natural selection processes to find the best match for a given file name through the following steps:

**Population Initialization:** Generates an initial population of random strings simulating potential file names.
**Selection:** Evaluates the population, selecting the best matches based on Levenshtein distance to the target file name.
**Crossover:** Combines pairs of selected individuals to produce offspring, ensuring a diverse gene pool.
**Mutation:** Randomly alters some characters in the offspring to introduce variations and avoid local optima.
**Iteration:** Repeats the selection, crossover, and mutation steps over several generations until the best match is found.

This approach significantly improves the chances of finding files with names that are not exactly known or have typographical errors.

## Libraries
**SpeechRecognition:** Library for performing speech recognition, with support for several engines and APIs, online and offline.
**pywin32:** Python extensions for Windows, providing access to many of the Windows APIs from Python.
