import re

class csvcell(object):

    def __init__(self, larray):

        self.larray = larray

    def ipv4_change(self,ccolumn,csvalue,cvalue):
        # change ipv4 address as specified
        try:
            c1 = self.larray[ccolumn].split('.')
            c1i = int(c1[csvalue]) + cvalue
            c1[csvalue] = str(c1i)
            self.larray[ccolumn] = '.'.join([str(elem) for elem in c1])
        except:
            pass
        return self.larray[ccolumn]

    def regexchg(self,ccolumn,csvalue,cvalue):
        # change value in specified column
        try:
            self.larray[ccolumn] = re.sub(csvalue,cvalue,self.larray[ccolumn])
        except:
            pass
        return self.larray[ccolumn]

    def regexsrcrep(self,scolumn, svalue, ccolumn, csvalue, cvalue):
        # search for a value in specified column if found then change value in 2nd specified column
        try:
            if re.search(svalue,self.larray[scolumn]):
                self.larray[ccolumn] = re.sub(csvalue,cvalue,self.larray[ccolumn])
        except:
            pass
        return self.larray[ccolumn]

    def regexsrc2rep(self,scolumn, svalue, s2column, s2value, ccolumn, csvalue, cvalue):
        # search for a value in specified two columns if found then change value in 2nd specified column
        try:
            if (re.search(svalue,self.larray[scolumn])) and (re.search(s2value,self.larray[s2column])) :
                self.larray[ccolumn] = re.sub(csvalue,cvalue,self.larray[ccolumn])
        except:
            pass
        return self.larray[ccolumn]