# Start with an empty dictionary
locker_dict = {}

# Fill the dictionary with key and values (closed)
for locker in range(1, 101):
    locker_dict[locker] = "closed"

# Loop through students
for student in range(1, 101):
    # Loop through lockers for every student
    for locker in range(student, 101, student): 
        if locker%student == 0: # Checking the locker number can be divided by student number. If so it is the student number's multiplier
            print(f"locker: {locker}, student: {student}, locker status (before open-close action): {locker_dict[locker]}") # Before open-close action, shows the status
            if locker_dict[locker] == "open":
                locker_dict[locker] = "closed" # Reverse the status
                print(f"locker: {locker}, student: {student}, locker status (after open-close action): {locker_dict[locker]}") # After open-close action, shows the status
            else:
                locker_dict[locker] = "open" # Reverse the status
                print(f"locker: {locker}, student: {student}, locker status (after close-open action): {locker_dict[locker]}") # After open-close action, shows the status

# Print the results
print(locker_dict)


# Print the results with row by row format
for key, value in locker_dict.items():
    print(f"{key}: {value}")
