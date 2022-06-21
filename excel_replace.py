import sys
import openpyxl
import json

def create_cfg_example():
    cfg = {"source_file" : "sample.xlsx", "target_file" : "target.xlsx", "sheet" : "Sheet2", "mode" : "basic/RE", \
            "text_original/pattern" : "", "text_new" : ""}

    with open("cfg.json.example", "w+") as f:
        json.dump(cfg, f, indent=4)

def main(src_file=None, dst_file=None):
    with open("cfg.json", "r") as f:
        cfg = json.load(f)
    (sheet_name, mode, text_old, text_new) = (cfg["sheet"], cfg["mode"], cfg["text_original/pattern"], cfg["text_new"])
    
    if ((src_file == None) or (dst_file == None)):
        (src_file, dst_file) = (cfg["source_file"], cfg["target_file"])


    wb = openpyxl.load_workbook(src_file)
    ws = wb[sheet_name]

    i = 0
    #o_string = "3.  Make toolbar icons customizable.\n"
    for r in range(1,ws.max_row+1):
        for c in range(1,ws.max_column+1):
            s = ws.cell(r,c).value
            if s != None and text_old in s: 
                ws.cell(r,c).value = s.replace(text_old,text_new) 

                print(f"row {r} col {c} updated")
                i += 1

    wb.save(dst_file)
    print(f"{i} cells updated")

if __name__ == '__main__':
    if len(sys.argv) >= 3:
        (src_file, dst_file) = sys.argv[-2:]
        main(src_file, dst_file)
    else:
        main()
        #create_cfg_example()