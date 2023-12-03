import datetime

def timestamp(): 
    # Get the current date and time
    now = datetime.datetime.now()

    return now.strftime("%d.%m.%y %H:%M:%S")

# Print the formatted date and time string
# print("Current date and time:", timestamp() )

