#!/bin/bash
###########################################################################
# Shell Script to backup Appliances SonicWALL firmware version 5.9 and above#
###########################################################################

# VARIABLE
#
# Do not remove the quotation marks :
#
# Directory for the log
DATE_TIME_UNDERLINE=$(date +%d"-"%m"-"%y"_"%H"-"%M)
DATE_TIME_PIPE=$(date +%d"-"%m"-"%y"|"%H"-"%M)
log="/var/log/sonicwall_bkp.log"
# User Sonicwall | IF YOU CHANGES IN THE NAME OF ADMIN SYSTEM> PREFERENCES CHANGE THE VARIABLE LOGIN
login="admin"
# Password sonicwall
password="PASSWORD"
# Address sonicwall | IP OR FQDN
host="192.168.168.168"


/usr/bin/expect <<EOF
spawn ssh $login@$host
expect -re ".*?assword:"
send "$password\n"
expect -re ">"
send "export current-config sonicos ftp ftp://USER_FTP:PASSWORD_FTP@IP_SERVER_FTP/bkp-SonicWall-$DATE_TIME_UNDERLINE.exp\n"

expect -re ">"
send "exit\n"
EOF

if [ $? = 0 ] ; then
	echo " $DATE_TIME_PIPE $host BKP done successfully!";
else 
	echo "$DATE_TIME_PIPE - $host BKP HAS NOT BEEN DONE!"
fi

done