#encoding=utf-8

import os

cmd="python3 ../export.py -f ./hero.xlsx -s 英雄 -t ./hero_data.py"
os.system(cmd)

cmd="python3 ../export.py -f ./hero.xlsx -s 全局参数 -t ./global_data.py -k global"
os.system(cmd)

import hero_data
heroData = hero_data.GetData()
print(heroData)

import global_data
globalData = global_data.GetData()
print(globalData)

cmd="python3 ../export.py -f ./hero.xlsx -s 英雄 -t ./hero_data.json"
os.system(cmd)

cmd="python3 ../export.py -f ./hero.xlsx -s 全局参数 -t ./global_data.json -k global"
os.system(cmd)

cmd="python3 ../export.py -f ./hero.xlsx -s 英雄 -t ./hero_data.lua"
os.system(cmd)

cmd="python3 ../export.py -f ./hero.xlsx -s 全局参数 -t ./global_data.lua -k global"
os.system(cmd)
