my_link = "K:\MainFolder\SubFolder1\SubFolder2"

find_character = my_link.find("\\")
rest_of_the_text = my_link[find_character:]
new_link = "servername" + rest_of_the_text

print("my_link >> ", my_link)
print("new_link >> ", new_link)
