import random

# Pre-defining users
user_dict = {1: "User 1"
             ,2: "User 2"
             ,3: "User 3"
             ,4: "User 4"
             ,5: "User 5"
             ,6: "User 6"
             ,7: "User 7"
             ,8: "User 8"
             ,9: "User 9"
             ,10: "User 10"
             }


# Finding the upper limit in the dictionary. We will be using this to calculate random value
dict_items_count = len(user_dict) # result => 10


selected_user_w_randrange = random.randrange(1,dict_items_count) # When we want to print both dictionary key and values it could calculated different random values. (This is not expected behavior.)
selected_user_w_randint = random.randint(1,dict_items_count)

print("Selected User using randrange >> ", str(selected_user_w_randrange) + " - " + user_dict[selected_user_w_randrange])
print("Selected User using randint >> ", str(selected_user_w_randint) + " - " + user_dict[selected_user_w_randint])
