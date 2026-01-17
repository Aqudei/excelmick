from tinydb import TinyDB, Query
from threading import Lock

_DB_CACHE: dict[str, TinyDB] = {}
_DB_LOCK = Lock()


def get_db(db_name: str) -> TinyDB:
    """
    Lazily create and cache TinyDB instances.
    One TinyDB per file path.
    """
    if db_name not in _DB_CACHE:
        with _DB_LOCK:  # thread-safe
            if db_name not in _DB_CACHE:
                _DB_CACHE[db_name] = TinyDB(f"./data/{db_name}.json")
    return _DB_CACHE[db_name]


# architects_db = TinyDB("./data/architects.json")
# surveyor_db = TinyDB("./data/surveyor.json")


def architect_find(key):
    db = get_db("architects")
    Architect = Query()
    return db.search(Architect.key == f"{key}")


def architect_set_status(data, status):
    db = get_db("architects")

    archs = architect_find(data.licence_number)

    Architect = Query()

    if archs and len(archs) > 0:
        db.update(
            {"status": status},
            Architect.key == f"{data.licence_number}",
        )
    else:
        db.insert({"status": status, "key": f"{data.licence_number}"})


def find_by_name(db_name: str, name):
    db = get_db(db_name)
    q = Query()
    result = db.search(q.name == name)
    if result and len(result) > 0:
        return result[0]

    return {}


def find_by_key(db_name: str, key):
    db = get_db(db_name)
    q = Query()
    result = db.search(q.key == key)
    if result and len(result) > 0:
        return result[0]

    return {}


def upsert(db_name: str, data, key):
    db = get_db(db_name)
    q = Query()
    items = db.search(q.key == key)
    if items and len(items) > 0:
        for item in items:
            doc_id = item.doc_id
            db.update(data, doc_ids=[doc_id])
    else:
        db.insert(data)
