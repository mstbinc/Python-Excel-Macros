import string

# The PFE alphabet excludes letters that can be confused as numbers
illegal_characters = ['I', 'O']

def get_alphabet():
	alpha = string.ascii_uppercase
	for c in illegal_characters:
		alpha = alpha.replace(c, '')
	return alpha

def number_to_letter(num):
	return num