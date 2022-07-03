class ProductInfoMessage:
    # name = ''
    # code = ''
    # supplier = ''
    # material = ''
    # price = 0

    p_name_line = '{}\r\n'
    p_code_line = 'Артикул магазина: {}\r\n'
    p_supplier_line = '\r\nПроизводитель: {}\r\n'
    p_material_line = 'Материал: {}\r\n'
    p_price_line = '\r\nЦена: {}\r\n\r\n'
    p_footer = ('Перед оформлением заказа, пожалуйста, прочитайте условия по ссылке в закреплённым посте.\r\n'
                'Оформить заказ или задать вопрос можно через бот @strollespbot')

    def __init__(self, code, name, price, supplier='', material='', ):
        self.name = name
        self.code = code
        self.supplier = supplier
        self.material = material
        self.price = price

    def __repr__(self):
        return (f'{__class__.__name__}('
                f'{self.code!r}, {self.name!r}, {self.price!r}, {self.supplier!r}, {self.material!r})')

    def print(self) -> str:
        return self.p_name_line.format(self.name) + \
               self.p_code_line.format(self.code) + \
               (self.p_supplier_line.format(self.supplier) if self.supplier else '') + \
               (self.p_material_line.format(self.material) if self.material else '') + \
               self.p_price_line + \
               self.p_footer
