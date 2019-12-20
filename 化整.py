from decimal import *
i=raw_input("any key to run or enter to quit")
print(i)
d=raw_input("kn=")
while i!="quit":
    x=input("input x=")
    x=float(x);
    d=float(d);
    y=int(x/d)
    #print(y)
    z=round(Decimal(x)-Decimal(d*y),5)
    if z > (0.5*d):
            r=d*(y+1)
            print(r)
    else:
        if z<(0.5*d):
            r=d*y
            print(r)
        else:
            if   (y%2==0):
                r=d*y
                print(r)
            else:
                r=d*(y+1)
                print(r)


  
