# Python code to
# demonstrate readlines()
import argparse
from cellobj import csvcell

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
                l1[ch['chgcol']] = pc.regexsrc2rep(ch['srccol'], ch['srcval'], ch['chgcol'], ch['chgsrcval'], ch['chgval'])
            else:
                pass
    except:
        pass
    return l1

def dup_line(line, file2):
    try:
        ch_list = [{'changetype': 'ipv4 change', 'chgcol': 0, 'srcval': 1, 'chgval': 0},
                   {'changetype': 'regexpsub', 'chgcol': 1, 'srcval': '255.255.252.0', 'chgval': '255.255.255.0'},
                   {'changetype': 'regexpsrcrep', 'srccol': 2, 'srcval': '^.*Floor1', 'chgcol': 4, 'chgval': r'Test - \g<0>', 'chgsrcval': '^.*$'}]
    # write original line to output
        newline = parse_line(line, ch_list)
        file2.writelines(newline)
        ch_list[0]['changeval'] = -1
        newline = parse_line(line, ch_list)
        file2.writelines(newline)
        ch_list[0]['changeval'] = 1
        newline = parse_line(line, ch_list)
        file2.writelines(newline)
    except:
        pass

def build_parser():
    cfg = argparse.ArgumentParser(
        description="Open a csv file and duplicate lines."
    )
    cfg.add_argument("path1", help="Input excel file")
    cfg.add_argument("path2", help="Output excel file")
    cfg.add_argument("-o", "--output-path", default="compared.xlsx",
                     help="Path of the comparison results")
    cfg.add_argument("-m", "--merge-path", default="merged.xlsx",
                     help="Path of the merged results")
    cfg.add_argument("--skiprows", help='number of rows to skip', type=int,
                     action='append', default=None)
    return cfg


def main():
    cfg = build_parser()
    opt = cfg.parse_args()
    # Using readlines()
    file1 = open(opt.path1, 'r')
    file2 = open(opt.path2, 'w')
    #treat header line differently
    newline = file1.readline()
    file2.writelines(newline)
    file1 = open(opt.path1, 'r')
    lines = file1.readlines()[1:]
    # interate through the rest of the lines
    for index, line in enumerate(lines):
        dup_line(line,file2)

    file1.close()
    file2.close()

if __name__ == '__main__':
    main()