# powerpoint_auto
## Overview 
This repo reads data from blockforce google sheets, and updates the performance update powerpoint. After updating powerpoint with the latest data it saves the ppt as images to then post on twitter 
# How to set up
## Disclaimers
This was supposed to be setup on a server, but due to using the pywin32 package (windows client) this cannot be ran on a linux server. 
Using an azure windows virtual machine would be difficult to debug and not free, so I thought having it setup on someones windows machine would be the best. 

## Set Up Environment Variables 
Ask Cyrus for the following variables \
INVESTOR_KEY\
INVESTOR_KEY_ID \
BFC_TWIT_API_KEY\
BFC_TWIT_API_SECRET_KEY\
BFC_TWIT_ACCESS_TOKEN\
BFC_TWIT_ACESS_TOKEN_SECRET

## Set up Virtual Enviornment 
run the following command
conda env create -f environment.yml
this will create the env used by the batch file, the environment is in python 3.6 for a reason

## Setup Windows Task Scheduler 

