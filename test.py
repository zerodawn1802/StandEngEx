import re

# Test input string
input_text = "_this_ _that_ _here_ _ _ _______ _something_ **bold** *italic* _underline_"

# Updated regex to match **bold**, *italic*, and _underline_, but not _ (one underscore) or continuous underscores
pattern = r'(\*\*[^*]+\*\*|\*[^*]+\*|_(?!_+$)[^_\s]+_)'

# Find all matches
matches = re.findall(pattern, input_text)

# Output matches
print(matches)
