from dataclasses import dataclass

from typing import Optional


@dataclass

class Product:

    """상품 정보를 담는 데이터 클래스"""
    name: str
    price: str
    type_name: str
    product_id: int
    type_id: int
    

    def __post_init__(self):

        """데이터 검증"""
        if not self.name.strip():

            raise ValueError("상품명은 필수입니다.")
        

        # 가격이 숫자인지 확인

        try:
            float(self.price.replace(',', ''))

        except ValueError:

            raise ValueError("가격은 숫자만 입력 가능합니다.")
    

    @property

    def formatted_price(self) -> str:

        """포맷된 가격 반환"""

        try:

            price_num = float(self.price.replace(',', ''))

            return f"{price_num:,.0f}"

        except ValueError:

            return self.price
    

    def to_dict(self) -> dict:

        """딕셔너리로 변환"""

        return {
            'name': self.name,
            'price': self.price,
            'product_id' : self.product_id,
            'type_id' : self.type_id,       
            'type_name': self.type_name,
        }
    

    @classmethod

    def from_dict(cls, data: dict) -> 'Product':

        """딕셔너리에서 생성"""

        return cls(

            name=data.get('name', ''),

            price=data.get('price', ''),

            product_id=data.get('product_id', ''),

            type_id=data.get('type_id', ''),

            type_name=data.get('type_name', ''),


        )

