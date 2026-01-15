from tinydb import TinyDB, Query

architects_db = TinyDB("./data/architects.json")

def architect_find(registration_no):
    Architect = Query()
    return architects_db.search(Architect.registration_no == f"{registration_no}")


def architect_set_status(data, status):
    archs = architect_find(data.licence_number)

    Architect = Query()

    if archs and len(archs) > 0:
        architects_db.update(
            {"status": status},
            Architect.registration_no == f"{data.licence_number}",
        )
    else:
        architects_db.insert(
            {"status": status, "registration_no": f"{data.licence_number}"}
        )
