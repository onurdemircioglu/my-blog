#!/bin/bash
xdotool search "Mozilla Firefox" windowactivate --sync # Terminal screen switchs to Firefox

xdotool click --repeat 2 1
xdotool sleep 0.1

#xdotool key --clearmodifiers "p" # If you want to run this code on main screen
sleep 0.5
for a in {1..20..1};do # Repeats the code 20 times
    xdotool key --clearmodifiers "alt+m" "alt+o" "alt+o" "alt+l" "alt+a"
sleep 0.5
done

sleep 0.5
xdotool key --clearmodifiers "o" # To go back the main screen.
