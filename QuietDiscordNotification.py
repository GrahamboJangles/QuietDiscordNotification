from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from pystray import Icon as icon, Menu as menu, MenuItem as item
from PIL import Image
import threading
import ctypes
import time

# Function to show the window
def show_window(icon, item):
    icon.stop()

# Function to quit the program
def quit_program(icon, item):
    icon.stop()
    
    # exit() only exits the current thread, so we have to kill all the other threads too.
    # exit()
    import psutil
    import os

    parent_pid = os.getpid()
    parent = psutil.Process(parent_pid)
    children = parent.children(recursive=True)
    for child in children:
        child.kill()
    parent.kill()

# Function to get the current volume of Discord
def get_discord_volume():
    # Get all audio sessions
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        # Check if this is the Discord session
        if session.Process and session.Process.name() == "Discord.exe":
            # Return the current volume of Discord
            return session.SimpleAudioVolume.GetMasterVolume()
    return None

# Function to set the volume of Discord
def set_discord_volume(volume):
    global discord_volume
    
    # Set the global volume variable to the specified value
    discord_volume = volume
    
    # Get all audio sessions
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        # Check if this is the Discord session
        if session.Process and session.Process.name() == "Discord.exe":
            # Set the volume of Discord
            session.SimpleAudioVolume.SetMasterVolume(volume, None)

# Function to show the current volume of Discord
def show_volume(icon, item):
    # Get the current volume of Discord
    volume = get_discord_volume()
    if volume is not None:
        # Show a message box with the current volume of Discord
        ctypes.windll.user32.MessageBoxW(0, f"The current volume of Discord is {volume:.0%}", "Discord Volume", 0)

# Function to set the volume of Discord to a specific value
def set_volume(icon, item, volume):
    set_discord_volume(volume)

# Create an image for the system tray icon
image = Image.open("discord.ico")

# Create a menu for the system tray icon
menu = menu(
    # item('Show Volume', show_volume),
    # menu.SEPARATOR,
    item('Set Volume to 100%', lambda icon, item: set_volume(icon, item, 1.0)),
    item('Set Volume to 75%', lambda icon, item: set_volume(icon, item, 0.75)),
    item('Set Volume to 50%', lambda icon, item: set_volume(icon, item, 0.5)),
    item('Set Volume to 25%', lambda icon, item: set_volume(icon, item, 0.25)),
    item('Set Volume to 0%', lambda icon, item: set_volume(icon, item, 0.0)),
    menu.SEPARATOR,
    # item('Show', show_window),
    item('Quit', quit_program)
)

# Create a system tray icon with the specified image and menu
tray_icon = icon("QuietDiscordNotification", image, "QuietDiscordNotification", menu)

# Function to run the system tray icon in a separate thread
def run_tray_icon():
    tray_icon.run()

# Create a thread for the system tray icon
tray_thread = threading.Thread(target=run_tray_icon)

# Start the thread for the system tray icon
tray_thread.start()

# Set the interval (in seconds) for setting the volume
interval = 10

# Set the initial volume for Discord (0.0 to 1.0)
discord_volume = 0.1

# Make it run in the background
ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

while True:
    # Set the volume of Discord
    set_discord_volume(discord_volume)
    
    # Wait for the next interval
    time.sleep(interval)
