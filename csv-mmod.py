# Python code to
# change csv file as specified in ch_list
import argparse
from cellobj import csvcell
import pandas as pd

def parse_line(line, chlist):
    l1 = line.split(',')
    try:
        l1 = parse_column(l1,chlist)
        newline = ','.join([str(elem) for elem in l1])
    except:
        newline = line
    return newline

def parse_column(l1, chl):
    try:
        # change IP address by adding change
        pc = csvcell(l1)
        for ch in chl:
            if ch['changetype'] == 'ipv4 change':
                l1[ch['chgcol']] = pc.ipv4_change(ch['chgcol'],ch['srcval'],ch['chgval'])
            elif ch['changetype'] == 'regexpsub':
                l1[ch['chgcol']] = pc.regexchg(ch['chgcol'], ch['srcval'], ch['chgval'])
            elif ch['changetype'] == 'regexpsrcrep':
                l1[ch['chgcol']] = pc.regexsrcrep(ch['srccol'], ch['srcval'], ch['chgcol'], ch['chgsrcval'], ch['chgval'])
            elif ch['changetype'] == 'regexpsrc2rep':
                l1[ch['chgcol']] = pc.regexsrc2rep(ch['srccol'], ch['srcval'], ch['src2col'], ch['src2val'], ch['chgcol'], ch['chgsrcval'], ch['chgval'])
            else:
                pass
    except:
        pass
    return l1

def change_line(line, file2, file3):
    try:
        ch_list = [{'changetype': 'ipv4 change', 'chgcol': 0, 'srcval': 1, 'chgval': 0},
                   {'changetype': 'regexpsub', 'chgcol': 1, 'srcval': '255.255.252.0', 'chgval': '255.255.255.0'},
                   {'changetype': 'regexpsrcrep', 'srccol': 2, 'srcval': '^.*Floor3', 'chgcol': 4, 'chgval': r'Test - \g<0>', 'chgsrcval': '^.*$'},
                   {'changetype': 'regexpsrc2rep', 'srccol': 2, 'srcval': r'^LLB', 'src2col': 2, 'src2val': '^.*Floor2', 'chgcol': 4,
                    'chgval': r'Test2 - \g<0> Done', 'chgsrcval': '^.*$'}]
    # write original line to output
        ch_list[0]['changeval'] = 0
        newline = parse_line(line, ch_list)
        file3.writelines(newline)
    except:
        pass

def build_parser():
    cfg = argparse.ArgumentParser(
        description="Open a csv file and modify lines."
    )
    cfg.add_argument("path1", help="Input data excel file")
    cfg.add_argument("path2", help="Input changes excel file")
    cfg.add_argument("path3", help="Output modified excel file")
    cfg.add_argument("--skiprows", help='number of rows to skip', type=int,
                     action='append', default=1)
    return cfg


def main():
    cfg = build_parser()
    opt = cfg.parse_args()
    # Using readlines()
    odata=pd.read_csv(opt.path1)
    moddata=pd.read_excel(opt.path2)
    #treat header line differently
    mdata = moddata.merge(odata, on='Device Name', how='left')
    #file1 = open(opt.path3, 'r')
    # interate through the rest of the lines
    mdata.to_csv('phones-2nd.csv')
    print('Done with merge.')

if __name__ == '__main__':
    main()