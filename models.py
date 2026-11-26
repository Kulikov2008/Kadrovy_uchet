from dataclasses import dataclass
from typing import Optional

@dataclass
class Department:
    id: Optional[int] = None
    name: str = ""
    description: Optional[str] = None

@dataclass
class Employee:
    id: Optional[int] = None
    full_name: str = ""
    position: str = ""
    department_id: Optional[int] = None
    salary: float = 0.0

@dataclass
class Order:
    id: Optional[int] = None
    employee_id: Optional[int] = None
    order_type: str = ""
    details: Optional[str] = None