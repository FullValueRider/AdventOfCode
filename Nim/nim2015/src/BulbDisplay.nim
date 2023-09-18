type 
    Area* = tuple
        OrgX : int,
        OrgY : int,
        MaxX : int,
        MaxY : int       


    BulbDisplay* = object
        Array           : seq[seq[int]]
        UserBrightness  : bool


proc `UseBrightness=`*( me: var BulbDisplay, ipValue : bool) =
    me.UseBrightness = ipValue


proc UseBrightness*(me:BulbDisplay) : bool =
    return me.UseBrightness


proc SwitchOn*( me: var BulbDisplay, ipArea : Area) =
    # iparea is a variant containg array(0 to 3) of x1,y1 x2,y2
    for myX in ipArea.OrgX..ipArea.MaxX:
        for myY in ipArea.OrgY..ipArea.MaxY:
            if me.UseBrightness == true:
                me.Array[myX][ myY] +=  1   
            else:
                me.Array[myX][ myY]  = 1
                

proc SwitchOff*(me: var BulbDisplay, ipArea : Area) =
# iparea is a variant containg array(0 to 3) of x1, y1 x2,y2
    for myX in ipArea.OrgX..ipArea.MaxX:
        for myY in ipArea.OrgY..ipArea.MaxY:
            if me.UseBrightness:
                me.Array[myX][ myY] += - 1
                if me.Array[myX][ myY]  < 0: 
                    me.Array[myX][ myY] = 0
            else:
                me.Array[myX][ myY] = 0


proc Toggle*( me: var BulbDisplay, ipArea : Area) =
# iparea is a variant containg array(0 to 3) of x1, y1 x2,y2
    for myX in ipArea.OrgX..ipArea.MaxX:
        for myY in ipArea.OrgY..ipArea.MaxY:
            if me.UseBrightness:
                me.Array[myX][ myY] +=  2  
            else:
                me.Array[myX][ myY] =  if me.Array[myX][ myY] == 0:  1 else: 0


proc LitBulbs*(me: var BulbDisplay, ipArea : Area) : int =
    for myX in ipArea.OrgX..ipArea.MaxX:
        for myY in ipArea.OrgY..ipArea.MaxY:
            if me.Array[myX][ myY] > 0: 
                result += 1


proc Brightness*( me: var BulbDisplay, ipArea : Area) : int =
    for myX in ipArea.OrgX..ipArea.MaxX:
        for myY in ipArea.OrgY..ipArea.MaxY:
            result +=  me.Array[myX][ myY]

proc initBulbDisplay*( me: var BulbDisplay, ipxsize: int, ipysize:int, ipbrightness:bool): BulbDisplay =
    me.Array= newSeq[newSeq[int](ipySize)](ipysize)
    me.UserBrightness = ipbrightness