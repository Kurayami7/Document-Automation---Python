"""
Document Automation
Author: Areaf
Assignment 1
Date: 21 June, 2023
 Function to read the file and count word occurrences """
def openToCount(fileName):
    word_count = {}
    with open('sample.txt', 'r') as file:
        words = file.read().split() # Reading the file and basically using the split method to split the string of words into substrings
        
        # If the word is in the empty dictionary, then increase it's count. Otherwise, set it to 1
        for word in words:
            if word in word_count:
                word_count[word] += 1
            else:
                word_count[word] = 1
    
    file.close()
    return word_count

# Function to save the word frequency count to a file
def saveCount(word_count, fileName):
    with open('output.txt', 'w') as file:
        # Sort the word counts in descending order
        countSorted = sorted(word_count.items(), key=lambda x: x[1], reverse=True) # A sorted function. The .items() returns the word_count dict as a list of tuples, where each tuple contains a word and it's count ('word', count)
        # The key=lambda x: x[1] uses a lambda function, x is each tuple from the items list and x[1] is the second element of the tuple, which is the count.
        # reverse = True is used to sort in descending order. It's false by default, which is ascending order.
        
        # Tuple unpacking loop>
        for word, count in countSorted:
            file.write(f"{word}: {count}\n")

word_count = openToCount('sample.txt')
saveCount(word_count, 'output.txt')

print("Word frequency report saved successfully.")
