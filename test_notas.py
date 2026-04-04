import sys, os
sys.path.insert(0, os.path.abspath('src'))
os.environ["REGISTRODOC_MASTER_SALT"] = "test"
from rdsecurity import validar_nota_meduca

# Test Boolean
b_res = validar_nota_meduca(True)
print(f'Boolean True: {b_res}')
b_res2 = validar_nota_meduca(False)
print(f'Boolean False: {b_res2}')

# Test Range
r1_res = validar_nota_meduca('0.9')
print(f'Under: {r1_res}')
r2_res = validar_nota_meduca('5.1')
print(f'Over: {r2_res}')
