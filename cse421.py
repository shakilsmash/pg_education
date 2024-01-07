import os
import discord
import re
import pandas as pd
from dotenv import load_dotenv
import cse421_gsheet_v2 as gsheet
import cse421_gcal as gcal
from datetime import datetime, timedelta
from discord.utils import get
from discord.ext import commands
from openpyxl import load_workbook

workbook = load_workbook('input.xlsx')
sheet = workbook.active

SERVER = '421'

ENV_TOKEN = SERVER+'_TOKEN'
ENV_GUILD = SERVER+'_GUILD'
ENV_BOT = SERVER+'_BOT'

load_dotenv()
TOKEN = os.getenv(ENV_TOKEN)
GUILD = str(os.getenv(ENV_GUILD)).split("/")[0]
GUILD_ID = int(str(os.getenv(ENV_GUILD)).split("/")[1])
intents = discord.Intents.all()
client = discord.Client(intents=intents)
bot_id = os.getenv(ENV_BOT)

#ARF,SKZ,SRJ,S06,S34, S36, S49, S60
admin_list = ['711193765578276864','734061644845678693','734358798667350016','820277759918080060','601824072569061377','742826509823246340','940584874253893643','544905137550917634']

# ignore_list = ['shakilsmash#1900','CSE421 Bot#3769']-
            
@client.event
async def on_ready():
    global guild
    print('Bot logged in...')
    guild = get_guild()

@client.event
async def on_message(message):
    global guild
    user = message.author #vyrevy
    member = guild.get_member(user.id)
    msg = str(message.content)
    # print(msg)
    # print(member.id)
    if(member.id == bot_id):
        return
    
# DEBUGGER ==============================================================================
    if msg == 'ping':
        await message.channel.send('pong')   

# ADMIN ==============================================================================
    elif msg.startswith('_') and message.guild == None and member.id in admin_list:
        await member.send("Fetching marks of **" + msg[1:]+"**....\n== THEORY ==")
        await member.send(content=gsheet.getAllMarks(studentID=msg[1:]))
        await member.send("\n== LAB ==")
        await member.send(content=gsheet.getLabMarks(studentID=msg[1:]))

# FETCH ALL MARKS ==============================================================================
    elif message.content == '!marks' and message.guild == None:
        await member.send("No marks are released yet...")        
        return
        await member.send("Fetching your marks ["+member.nick[0:8]+"]...*Sorry if the response is slow... try again if I don't reply in few minutes.*")        
        await member.send(content=gsheet.getAllMarks_v3(studentID=int(member.nick.split('_')[0]),section=int(member.nick.split('_')[1])))

# FETCH MIDTERM MARKS ==============================================================================
    elif message.content == '!midmarks' and message.guild == None:
        await member.send("No marks are released yet...")        
        return
        await member.send("Fetching your midterm marks ["+member.nick[0:8]+"]...*Sorry if the response is slow... try again if I don't reply in few minutes.*")
        await member.send(content=gsheet.getMidMarks(studentID=member.nick[0:8]))

# FETCH FINAL MARKS ==============================================================================
    elif message.content == '!finmarks' and message.guild == None:
        await member.send("No marks are released yet...")        
        return
        await member.send("Fetching your midterm marks ["+member.nick[0:8]+"]...*Sorry if the response is slow... try again if I don't reply in few minutes.*")
        await member.send(content=gsheet.getFinMarks(studentID=member.nick[0:8]))

    # elif message.content == '!marks' and message.guild == None:
    #     await member.send("No marks are released yet...")        
    #     return
    #     await member.send("Fetching your marks...")        
    #     await member.send(content=gsheet.getAllMarks(discordID=member.id))
    
# FETCH ALL LAB MARKS ==============================================================================
    elif message.content == '!labmarks' and message.guild == None:
        await member.send("No marks are released yet...")        
        return
        for x in member.roles:
            if x.name.startswith("lab"):
                await member.send("Fetching your lab marks...")        
                await member.send(content=gsheet.getLabMarks(studentID=member.nick[0:8],section=x.name.split("-")[1]))      
        # await member.send("Fetching your lab marks...")        
        # await member.send(content=gsheet.getLabMarks(studentID=member.nick[0:8],section=member.nick[9:10]))

# FETCH ATTENDANCE ==============================================================================
    elif message.content == '!attendance' and message.guild == None:
        await member.send("No marks are released yet...")        
        return
        await member.send("Fetching your attendance...")        
        await member.send(content=gsheet.getAttendance(studentID=int(member.nick.split('_')[0]),section=int(member.nick.split('_')[1])))
# FETCH LAB MARKS ==============================================================================
    # elif message.content == '!labmarks':
    #     # user = message.author #vyrevy
    #     # print(message)
    #     await user.create_dm()
    #     await user.dm_channel.send(
    #         "Creating your marksheet...."\
    #         # +str(message.created_at)
    #     )
    #     await user.dm_channel.send(
    #         str(sheet.getLabMarks(user.id))
    #     )

# FETCH ANNOUNCEMENT (X DAYS) ==============================================================================
    elif msg.startswith('!announcement') and user.id != bot_id and message.guild == None:
        # user = message.author #vyrevy
        await member.send("Coming soon...")        
        return
        size = 7
        if msg.count('_')==2:
            size = int(msg.split('_')[-1])

        if('lab' in msg):
            messages = await client.get_channel(980325081614008330).history(limit=(size*3)).flatten()
        # elif('st' in msg):
        #     messages = await client.get_channel(937431655881248859).history(limit=(size*3)).flatten()
        else:
            messages = await client.get_channel(980325081362362387).history(limit=(size*3)).flatten()
        
        strr=""
        ret_str = list()
        start_date = datetime.today() - timedelta(days=(size+1))
        for message in messages:
            message_date = message.created_at
            if(start_date < message_date):
                ret_str.append("*"+str(message.author)+"* : **"+gcal.reformat(str(message_date))+"**\n"+message.content+"\n\n..")

        for i in range(len(ret_str)-1,-1,-1):
            if(len(ret_str[i])>0):
                await member.send(content=ret_str[i])
            else:
                await member.send(content="No announcements in the range given")

# FETCH ANNOUNCEMENT (X DAYS) ==============================================================================
    elif msg.startswith('!deadline') and user.id != bot_id and message.guild == None:
        await member.send("Construction ongoing....")        
        return
        send= list()
        
        if(msg.endswith('line')):
            send = gcal.fetch()
        elif(msg.endswith('d')):
            send = gcal.fetch(no_days=int(msg[10:-1]))
        elif(msg.endswith('e')):
            send = gcal.fetch(no_events=int(msg[10:-1]))
        else:
            send = "Hi! Looks like you tried to use the command **!deadline_Xe** or **!deadline_Xd**. Substitute the X with any number you want like: **!deadline_7d**"
            
        for i in range(0,len(send)):
            if(len(send[i])>0):
                await member.send(content=send[i])
            else:
                await member.send(content="No deadlines in the range given")
        
        
# VERIFICATION STEP 1: Student Existence ==============================================================================
    elif re.search("[0-9]{8}(-)[A-Z]{6}", msg) and str(member.id) != str(bot_id) and (message.guild == None or message.channel.id == 1023627273829625878):
        verification_code = str(re.findall("[0-9]{8}-[A-Z]{6}", msg)[0])
        print("\nstarted",verification_code,member.name)
        verification = gsheet.verify(verification_code)
        send = "Your verification process was **not successful**...A possible reason for this could be:\n"\
               "1. You are not verified on the CSE advising server. Please get the verification code from there first.\n"\
               "2. You have an updated student ID. Email **shafqat.hasan@bracu.ac.bd** from your G-Suite with your correct ID.\n"
        if(len(verification)==1):
            send = "Looks like you are already verified on the server! If this is false or you are changing your account, please email **shafqat.hasan@bracu.ac.bd** mentioning your old and updated Discord user name/ID."
        elif(len(verification)==4):
            # print(str(verification[0]),str(verification[3])[0:8],verification[1],verification[2])
            send = "You are now verified for CSE421 Fall 2023 Server...\n"\
                "**Student ID:** "+str(verification[3])[0:8]+"\n"\
                "**Nickname Changed to:** "+verification[3]+"\n"\
                "**Theory Section:** "+verification[1]+"\n"\
                "**Lab Section:** "+verification[2]+"\n"
            role = list()
            role.append(discord.utils.get(guild.roles, name=verification[0]))
            lab_role = "lab-"+str(verification[2]).zfill(2)
            role.append(discord.utils.get(guild.roles, name=lab_role))
            role.append(discord.utils.get(guild.roles, name="Student"))
            # print(role)
            await member.edit(nick=verification[3][0:31])
            await member.add_roles(*role)
            print("Verified:",member.name)
        try:
            await member.send(content=send)
        except:
            pass
    # SEND DISCORD ID ==============================================================================        
    elif msg == '!myid' and message.guild == None:
        send = "Your Discord User ID is: **"+str(member.id)+"**"
        await member.send(content=send)   
    
    elif msg == '!reverify':
        for memberr in guild.members:
            roles = memberr.roles
            if len(roles) > 1:
                role = list()
                role.append(discord.utils.get(guild.roles, name="Student"))
                await memberr.add_roles(*role)
                print(memberr.name)
    
    elif msg == '!give_lab':
        for memberr in guild.members:
            roles = memberr.roles
            print("Assigning:",memberr.name,len(roles))
            if (len(roles) == 3):
                role = list()
                print(memberr.roles)
                lab_ = gsheet.assign_lab(discordID=memberr.id)
                lab_role = "lab-"+str(lab_[0]).zfill(2)
                role.append(discord.utils.get(guild.roles, name=lab_role))
                await memberr.add_roles(*role)
                print("Assigned:",memberr.name,"\n")
                # break

@client.event
async def on_member_join(user):
    send = "Welcome to the CSE421 Fall 2023 Discord Server!\n"\
           "Please write/copy your verification code of your advising server and send it to me to get verified!\n"\
           "The verification code looks like: ***12345678-ABCDEF***. *i.e. 8 digit number,then a hyphen, then a 5 character string.*\n"\
           "If you do not have the verification code, please fill up this form (https://forms.gle/5HEkS8DwR6sGaCjx8) and a code will be emailed shortly (in 2 hours or so)."
    await user.send(content=send)
        
def get_guild():
    return client.get_guild(GUILD_ID)

# def help_text(index):
#     help_text = list()
    
#     #index = 0
#     help_text.append("Your verification process was **not successful**...A possible reason for this could be:\n"\
#                     "1. You did not fill up the google form (https://forms.gle/Fi1LQ1NaYgHEvkB5A)\n"\
#                     "2. You gave the wrong student ID or Discord User ID. Email **anis.sharif@bracu.ac.bd** from your G-Suite with your correct IDs.\n"\
#             "Your correct Discord User ID is **"+str(user.id)+"**. \n")
#     return help_text[index]
        
client.run(TOKEN)