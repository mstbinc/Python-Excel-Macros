import string

# The PFE alphabet excludes letters that can be confused as numbers
illegal_characters = ['I', 'O']

def get_alphabet():
	alpha = string.ascii_uppercase
	for c in illegal_characters:
		alpha = alpha.replace(c, '')
	return alpha

"""
TODO:
There is an error with this:
	f(48) should return 'AZ', but it retuns 'BZ'
"""
def number_to_letter(num):
	alpha = get_alphabet()
	base = len(alpha)
	string = ''
	while num != 0:
		if num == base:
			return index_to_char(base) + string
		remainder = num % base
		string = index_to_char(remainder) + string
		num = num / base
	return string

def letter_to_number(letters):
	letters = letters.upper()
	alpha = get_alphabet()
	base = len(alpha)
	position = len(letters) - 1
	total = 0

	for digit in letters:
		index = char_to_index(digit)
		place = (base ** position) * index
		total += place
		position -= 1

	return total

def char_to_index(char):
	return get_alphabet().find(char) + 1

def index_to_char(index):
	return get_alphabet()[index - 1]

# debug
if __name__ == "__main__":

	alpha = get_alphabet()
	print ''
	print alpha
	print "length: {}".format(len(alpha))
	print ''

	testnum = 24
	print "{} converted to letter: {}".format(testnum, number_to_letter(testnum))

	testletters = 'z'
	print "{} converted to number: {}".format(testletters, letter_to_number(testletters))