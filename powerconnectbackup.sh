#!/usr/bin/expect

# Usage:

# powerconnectbackup.sh <tftp server> <switch IP> <username> <password> <enable password if needed>

set timeout 60
set tftp [lindex $argv 0]
set name [lindex $argv 1]
set user [lindex $argv 2]
set password [lindex $argv 3]
set enablepw [lindex $argv 4]
set index 0
set transfer 0

if { [ string length $enablepw ] == 0 } {
    set enablepw $password
}

send_user "Opening connection to $name\n"
spawn telnet $name 

expect {
    User {
        if {$index == 1} {
            send_user "Incorrect username or passwordi\n"
            exit
        }
        send "$user\n"
        incr index
        exp_continue
    }
    Password {  
        send "$password\n"
        exp_continue
    }
    "#" {
        if {$transfer == 1} {
            send_user "Transfer complete\n"
            exit
        }
        send "copy startup-config tftp://$tftp/$name\n"
        incr transfer
        exp_continue
    }
    "Are you sure you want to start? (y/n)" {
        send "y\n"
        exp_continue
    }
    ">" {
        send "enable\n"
        exp_continue
    }
    "successfully" {
    }
}
send_user "Closing connection to $name\n"

close $spawn_id