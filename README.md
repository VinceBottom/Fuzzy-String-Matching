This will allow you to match two datasets based on names of varying spellings and styles (ex: one dataset reads "Phil R. Smith" and one reads "Phil Smith") along with date of birth in order to find potential matches.

It uses Python's difflib library to calculate Levenshtein distance between two strings, and where it meets the criteria (cutoff value written in the script) it will find the corresponding entry, 
so long as the date of birth matches. It would need to be modified to catch errors in date of birth (error of off by one day, etc.). For my purposes, this was not relevant, but it may be to you.
