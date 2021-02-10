import winapps
import subprocess
import socket






# get each application with list_installed()

def get_serialnumber():
    Data = subprocess.check_output(["wmic", "bios", "get", "serialnumber"])
    serial_number = str(Data)
    removal_element = [r"\r", r"\n", "b'", "'"]
    array_length = len(removal_element)
    for i in range(array_length):
        serial_number = serial_number.replace(r"{}".format(removal_element[i]), '')
    return(serial_number.strip())
    
def host_name():
    Hostname=socket.gethostname()
    return(Hostname)


f = open("data.txt", "w")
for item in winapps.list_installed():
    f.write("Application name:{} \t Version:{} \n".format(item.name,item.version))
f.write("{} \n {}".format(get_serialnumber(),host_name()))
f.close()











