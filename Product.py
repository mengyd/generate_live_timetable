class Product:
    def __init__(self, id, name = None, alias = None, link = None, \
        mespace_id = None, codes = [], product_type = None):
        self.id = id
        self.name = name
        self.alias = alias
        self.link = link
        self.mespace_id = mespace_id
        self.codes = codes
        self.product_type = product_type
        