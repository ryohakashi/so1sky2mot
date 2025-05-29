import os


def listFilesInbox(strInbox):    
    listFiles = os.listdir(strInbox)
    return listFiles

def moveFilesBackup(strInbox, strBackup, listFiles: []):
    for filename in listFiles:
        if (filename.endswith(".xls")):
            # print(os.path.join(strbackup, filename))
            os.rename(os.path.join(strInbox, filename), os.path.join(strBackup, filename))


if __name__ == "__main__":
    strInbox = "./inbox"
    strBackup = "./backup"

    FilesInbox = listFilesInbox(strInbox)
    moveFilesBackup(strInbox, strBackup, FilesInbox)