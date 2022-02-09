import pandas  as pd
d = {'Computer':1500,'Monitor':300,'Printer':150,'Desk':250}
print(d.keys())
temp = pd.DataFrame(d,index=['3'])
print(temp)