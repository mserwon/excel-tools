from cellobj import csvcell

line = "10.129.24.0,255.255.255.0,CFC-LubbockAnnex-Floor3,Lubbock Annex,FALSE,FALSE"
l1 = line.split(',')
column = 0
svalue = 1
cvalue = -1

nc = csvcell(l1)

nc2 = nc.ipv4_change(column,svalue,cvalue)
