#!/usr/bin/env python3

import subprocess
import sys
from datetime import datetime 
from os import path
from O365 import Account
from O365 import Message
from pathlib import Path
from datetime import date

#Import config 
from input_config import voltage_brownout


def main():

    #Define Variables
    final = ""

    status_read = status = ""
    time_read = time_elapsed = 0
    charge = runtime = voltage_in = 0
    
    path = "//home/pi/Documents/telegraf-nut-input-master/"
    
    #Index 0 is time, 1 is battery, 2 is state
    data = read_data(path)
    previous_time = data[0]
    previous_charge = data[1]
    previous_state = data[2]
    
    # If an argument isn't supplied, exit
    if len(sys.argv) <= 1:
        print('Please include a valid NUT UPS name\n  Example: input.py ups_name@localhost')
        sys.exit(1)

    # Set the UPS name from an argument and host
    full_name = sys.argv[1]
    ups_name = full_name.split('@')[0]

    # Get the data from upsc
    data = subprocess.run(["upsc", full_name], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)

    # For each line in the standard output
    for line in data.stdout.splitlines():
        # Replace ": " with a ":", then separate based on the colon
        line = line.replace(': ', ':')
        key = line.split(':')[0]
        value = line.split(':')[1]
        
        #Get variables for later use
        if(key == "ups.status"):
            status = value
        elif(key == "battery.charge"):
            charge = value
        elif(key == "input.voltage"):
            voltage_in = value
        elif(key == "battery.runtime"):
            runtime = value
            
        try:
            # If the value is a float, ok
            float(value)
        except ValueError:
            # If the value is not a float (i.e., a string), then wrap it in quotes (this is needed for Influx's line protocol)
            value = f'"{value}"'
            
        # Create a single data point, then append that data point to the string
        data_point = f"{key}={value},"
        
        final += data_point

    # Format is "measurment tag field(s)", stripping off the final comma
    print("ups,"+"ups_name="+ups_name, final.rstrip(','))

    """Send message if certain conditions are met (flag):
    0 = normal
    1 = blackout
    2 = brownout
    3 = surge"""
    state = 0
    
    #Determine State
    if(status == "OB" or status == "OB DISCHRG"):
        state = 1
    elif(brownout_check(voltage_in)):
        state = 2
    else:
        state = 0
        
    state = 1
    
    
    #Send email if UPS state changes (and at 60/40/20%)
    if(int(previous_state) != state):
        send_email(status, charge, runtime, voltage_in, state)
    elif(state != 0 and int(previous_charge) > 60 and charge <= 60):
        send_email(status, charge, runtime, voltage_in, state)
    elif(state != 0 and int(previous_charge) > 40 and charge <= 40):
        send_email(status, charge, runtime, voltage_in, state)
    elif(state != 0 and int(previous_charge) > 20 and charge <= 20):
        send_email(status, charge, runtime, voltage_in, state)


def read_data(path):
    
    #Open and read data file
    body_file = open(path + "data.txt", "r")
    body_txt = body_file.readlines()
    data = [body_txt[0][:-1], body_txt[1][:-1], body_txt[2]]
    body_file.close()
    
    return data


def send_email(status, charge, runtime, voltage_in, state):
    
    #Define Variables
    account = 0
    m = 0
    
    today = date.today()
    
    #Authenticate Process
    account = authenticate()
    
    #Set the Sender of Message
    m = set_sender(account, 'power-status-monitor@acdsystems.com')
    
    #Create the Message
    m = create_message(m, 'julian.france@acdsee.com', status, charge, runtime, voltage_in, state)
    
    #Save and send Message
    m.save_message()
    m.send()
    
    #Create a Log of Email
    create_log(m, today)
    
    #Get current time
    today = datetime.now()
    time_now = today.strftime("%H%M%S")

    #Change time + state of previous sent email
    body_file = open("//home/pi/Documents/telegraf-nut-input-master/data.txt", "w")
    body_file.write(time_now + "\n" + str(charge) + "\n" + str(state))
    body_file.close()
    
    
def authenticate():

    credentials = ('9e8ff30d-6732-4bbb-bfad-774e6e4b7514', 'Wx37Q~j_v7EXyx2Blb2bGqSkuKcFV_sqfuRWK')

    account = Account(credentials, auth_flow_type='credentials', tenant_id='d978af43-6932-48f7-9d42-dd1af9cb0251')
    account.authenticate()
        #print('Authenticated!')
    
    return account


def set_sender(account, sender):
    m = account.new_message(resource=sender)
    return m


def create_message(m, recipient, status, charge, runtime, voltage_in, state):

    #Create the message in system
    m.to.add(recipient)
    
    #Edit the subject line to include simple status warning
    option1 = ""
    option2 = ""
    option3 = ""
    if(status == "OL"):
        option1 = ""
        option2 = "UPS is currently on line power and fully charged."
    elif(status == "OL CHRG"):
        option1 = "[CHARGING]"
        option2 = "UPS is currently on line power and charging."
    elif(status == "OB" or status == "OB DISCHRG"):
        option1 = "[WARNING]"
        option2 = "UPS is on battery power!"
        
    #Add a line in the event of a brownout
    if(state == 2):
        if(status == "OL" or status == "OL CHRG"):
            option1 = "[WARNING]"
            option3 = "Brownout warning: UPS is still on line power but voltage has dropped. "
        elif(status == "OB" or status == "OB DISCHRG"):
            option3 = "Brownout warning: UPS has switched to battery due to low supply voltage. "
    
    subject_txt = ("UPS Report - Bear Mountain " + option1)
    m.subject = subject_txt
    
    #Main body of email
    body = """
        <html>
            <body>
                <strong>This is an automated message sent from Power Status Monitor at Bear Mountain.</strong>
                <p>
                   <p style="color:#996600;"> {option3} </p>
                    <br>{option2}<br>
                    <b>Current UPS info:</b><br>
                    Charge Level: {charge}%<br>
                    Runtime Remaining: {runtime}s<br>
                    Voltage In: {voltage_in}V<br><br>
                    For more information visit:
                    <a href="http://192.168.71.110:3000/d/7c7x6fZgz/ups-dash?orgId=1&refresh=10s"><br>Grafana Dashboard</a>
                </p>
            </body>
        </html>
        """.format(**locals())
    
    m.body = body
    m.attachments.add('//home/pi/Documents/telegraf-nut-input-master/attachment.txt')
   
    #print('Message sent.')
    return m
    
    
def create_log(m, today):
    log_count = 0
    date = today.strftime("_%b%d_%Y_")
    log_name = "email_log_"+str(log_count)+date+".eml"
    
    while(path.exists("email_log_"+str(log_count)+date+".eml")):
        log_count+=1
        log_name = "email_log_" + str(log_count)+date+".eml"
        
    m.save_as_eml(to_path = Path('//home/pi/Documents/telegraf-nut-input-master/Logs/' + log_name))


def brownout_check(voltage_in):
    #Threshold at which brownout is detected
    return(float(voltage_in) < voltage_brownout)


if __name__ == "__main__":
    main()