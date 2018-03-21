#if 用法
a=1
b=2
if a == 2:
    print('yyy')
elif a == b:
    print('sss')
else:
    print('nnn')


#宣告def 類似c#的void

def a(a):
    print(a)



a('abc')


#類別
class f1:
    def __init__(self,nu,sp):
        self.sp  = sp
        self.nu = nu

    def addsp(self):
        self.sp += 1
        print(self.sp)
    def show(self):
        print(self.nu)
        
    

dad = f1('dad',30)
mom = f1('mom',25)

dad.addsp()
mom.addsp()
dad.show()


    
    

        
