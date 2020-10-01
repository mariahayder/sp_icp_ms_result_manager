
def find_smpl_number(FILENAME):
    position = FILENAME.find('SMPL')
    SMPL = FILENAME[0]
    SMPL += FILENAME[1]
    SMPL += FILENAME[2]
    print(SMPL)
    return SMPL