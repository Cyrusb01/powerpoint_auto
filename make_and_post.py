"""
This script when run activates the virtual env, and runs the post_twitter.py file. 
This is the script run by the .bat file
"""
import subprocess

subprocess.run('conda activate powerpoint_3.6 && python post_twitter.py && conda deactivate', shell=True)
