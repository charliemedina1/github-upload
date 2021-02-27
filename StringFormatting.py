data = ("John", "Doe", 53.44)
format_string = "Hello"

print("%s" % format_string +" %s %s. Your current balance is $%s" % data)

data = ("John", "Doe", 53.44) #like a list so it is a string
format_string = "Hello %s %s. Your current balance is $%s."

print(format_string % data)
