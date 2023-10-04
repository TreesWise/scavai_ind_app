from fastapi import FastAPI, Depends, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from fastapi.security import OAuth2PasswordRequestForm
from datetime import timedelta
from fastapi.staticfiles import StaticFiles
import os
import io
# from config import ACCESS_TOKEN_EXPIRE_MINUTES
# from database_conn import database
# from helper import authenticate_user, create_access_token, get_current_active_user
from custom_data_type import  UserInput,DocDataModel

from subprocess import STDOUT, check_call , call,run

import torch

from app_func import predict,excel,pdf

#import torch

app = FastAPI()

app.mount("/files", StaticFiles(directory="files"), name="files")

def load_git():
   github='ultralytics/yolov5'
   torch.hub.list(github, trust_repo=True)
   model = torch.hub.load("ultralytics/yolov5", "custom", path = "./rings18.pt", force_reload=True)
   model.classes=[3 ,10,11 ,12, 17]
#     run(['apt-get', 'update']) 
#     run(['apt-get', 'install', '-y', 'libgl1'])
#     run(['apt-get', 'install', '-y', 'libglib2.0-0'])
#     run(['apt-get', 'install' ,'-y','abiword']) 
   # check_call(['apt-get', 'update'], stdout=open(os.devnull,'wb'), stderr=STDOUT)
   # check_call(['apt-get', 'install', '-y', 'libgl1'], stdout=open(os.devnull,'wb'), stderr=STDOUT)
    #check_call(['apt-get', 'install', '-y', 'libglib2.0-0'], stdout=open(os.devnull,'wb'), stderr=STDOUT)

   # check_call([ 'apt-get', 'update','-y'], stdout=open(os.devnull,'wb'), stderr=STDOUT)
   # check_call([ 'apt-get', 'install' ,'-y','abiword'], stdout=open(os.devnull,'wb'), stderr=STDOUT)
    
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
async def database_connect():
   
    run(['apt-get', 'update']) 
    run(['apt-get', 'install', '-y', 'libgl1'])
    run(['apt-get', 'install', '-y', 'libglib2.0-0'])
    run(['apt-get', 'install' ,'-y','abiword']) 
    #load_git()
    




@app.post("/predict")
async def fetch_data(userinput: UserInput):
    #print(userinput.dict())
    #runandget()
    run(['apt-get', 'update']) 
    run(['apt-get', 'install', '-y', 'libgl1'])
    run(['apt-get', 'install', '-y', 'libglib2.0-0'])
    run(['apt-get', 'install' ,'-y','abiword']) 
    pred = predict(userinput.dict())
    print(pred)
    return pred



@app.post("/excel")
async def fetch_excel(docInput: DocDataModel):
    #print(userinput.dict())
    #runandget()
    run(['apt-get', 'update']) 
    run(['apt-get', 'install', '-y', 'libgl1'])
    run(['apt-get', 'install', '-y', 'libglib2.0-0'])
    run(['apt-get', 'install' ,'-y','abiword']) 
    pred = excel(docInput.dict())
    print(pred)
    return pred

@app.post("/pdf")
async def fetch_excel(docInput: DocDataModel):
    #print(userinput.dict())
    #runandget()
    run(['apt-get', 'update']) 
    run(['apt-get', 'install', '-y', 'libgl1'])
    run(['apt-get', 'install', '-y', 'libglib2.0-0'])
    run(['apt-get', 'install' ,'-y','abiword']) 
    pred = pdf(docInput.dict())
    print(pred)
    return pred