
import datetime as dt
import json

def create_db_timeline(up_time_content=None):
    """
    DB Table의 col중 UP_TIME 같이 업데이트가 되는 테이블들은 해당 매서드로 업데이트 
    UPDATE 회수는 68번 제한 oracle varchar2()는 4000이 한계
    json.loads(obj)로 이용가능한 str return
    """


    # create 시 처음 json데이터 생번
    data=[]
    cols=['a','b','c']
    if up_time_content is None:
        rows=[]
        now = str(dt.datetime.now()).split('.')[0]
        rows.append('create')
        rows.append(str(0))
        rows.append(now)
        tmp = dict(zip(cols,rows))
        data.append(tmp)
        res = {
            'meta':{
                'desc':'data_update_timeline',

                'cols':{
                    'a':'db_method',
                    'b':'count',
                    'c':'timeline',

                }
            },
            'data':data
        }
        return json.dumps(res,ensure_ascii=False)
    # 데이터 생성이후 update시마다 계속 업데이트
    else : 
        load_json = json.loads(up_time_content)['data']
        update_count=len(load_json)
        if len(load_json) == 68:
            return up_time_content
        rows=[]
        now = str(dt.datetime.now()).split('.')[0]
        rows.append('update')
        rows.append(str(update_count))
        rows.append(now)
        tmp = dict(zip(cols,rows))
        load_json.append(tmp)
        res = {
            'meta':{
                'desc':'data_update_timeline',

                'cols':{
                    'a':'db_method',
                    'b':'count',
                    'c':'timeline',

                }
            },
            'data':load_json
        }
        
        return json.dumps(res,ensure_ascii=False)



def cal_std_day(befor_day):   
    x = dt.datetime.now() - dt.timedelta(befor_day)
    year = x.year
    month = x.month if x.month >= 10 else '0'+ str(x.month)
    day = x.day if x.day >= 10 else '0'+ str(x.day)  
    return str(year)+ '-' +str(month)+ '-' +str(day)