rem goto :a

netsh interface ip show config
netsh interface ip delete arpcache
netsh firewall set opmode disable

rem This section works with 2003 and XP
arp -s 172.30.50.140 00-01-00-00-04-00
arp -s 172.30.50.141 00-01-00-00-04-01
arp -s 172.30.50.142 00-01-00-00-04-02

rem this section works with Vista
netsh interface ip set address "Local Area Connection 12" static 172.30.50.1 255.255.0.0
netsh interface ip set address "Local Area Connection 13" static 172.31.50.1 255.255.0.0
netsh interface ip set address "Local Area Connection 14" static 172.32.50.1 255.255.0.0
netsh interface ip set address "Local Area Connection 15" static 172.33.50.1 255.255.0.0

netstat -ano

:a





netsh -c "interface ip" set neighbors "Local Area Connection 12" 172.30.50.140 00-01-00-00-04-00
netsh -c "interface ip" set neighbors "Local Area Connection 12" 172.30.50.141 00-01-00-00-04-01
netsh -c "interface ip" set neighbors "Local Area Connection 12" 172.30.50.142 00-01-00-00-04-02

netsh -c "interface ip" set neighbors "Local Area Connection 13" 172.31.50.143 00-01-00-00-04-03
netsh -c "interface ipv4" set neighbors "Local Area Connection 4" 172.31.50.160 00-01-00-00-06-01
netsh -c "interface ipv4" set neighbors "Local Area Connection 4" 172.31.50.161 00-01-00-00-06-02

goto :b
netsh -c "interface ipv4" set neighbors "Local Area Connection 5" 172.32.50.162 00-01-00-00-06-03
netsh -c "interface ipv4" set neighbors "Local Area Connection 5" 172.32.50.163 00-01-00-00-06-04

netsh -c "interface ipv4" set neighbors "Local Area Connection 5" 172.32.50.180 00-01-00-00-08-00
netsh -c "interface ipv4" set neighbors "Local Area Connection 6" 172.33.50.181 00-01-00-00-08-01
netsh -c "interface ipv4" set neighbors "Local Area Connection 6" 172.33.50.182 00-01-00-00-08-02
netsh -c "interface ipv4" set neighbors "Local Area Connection 6" 172.33.50.183 00-01-00-00-08-03

:b
pause
