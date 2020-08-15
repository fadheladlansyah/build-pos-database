import xlrd
import psycopg2
from sqlalchemy import create_engine, MetaData, Column, Integer, String
from sqlalchemy.orm import sessionmaker
from sqlalchemy.ext.declarative import declarative_base


# ---- create engine and table

conn_str = 'postgresql://{username}:{password}@localhost:5432/{database}'.format(
    username='postgres',
    password='rootpassword',
    database='postgres',
    )
engine = create_engine(conn_str)

meta = MetaData(bind=engine, reflect=True, schema="public")
Base = declarative_base(metadata=meta)

class PointofSales(Base):
    __tablename__ = 'yaks_hotel_pos201701'

    index = Column(Integer, primary_key=True)
    code = Column('Code', String(10))
    description = Column('Description', String(50))
    qty = Column('Qty', Integer)
    food_rev = Column('Food Rev', Integer)
    bev_rev = Column('Bev Rev', Integer)
    oth_fb_rev = Column('Oth FB Rev', Integer)
    amount = Column('Amount +/+', Integer)
    sv_chg = Column('SvChg', Integer)
    tax = Column('Tax', Integer)
    amount_net = Column('Amount Nett', Integer)
    cash = Column('Cash', Integer)
    credit_card = Column('Credit Card', Integer)
    city_ledger = Column('City Ledger', Integer)
    bill_fo = Column('Bill FO', Integer)
    other_pymt = Column('Other Pymt', Integer)
    remarks = Column('Remarks', String(20))
    bill_no = Column('Bill No', String(12))

Base.metadata.create_all(engine)


# ---- input data into table

workbook = xlrd.open_workbook("POS-Report-Detail Outlet-201701-clean.xlsx")
sheet = workbook.sheet_by_name("POS-Report-Detail Outlet-201701")

Session = sessionmaker(bind=engine)
session = Session()

for rowx in range(1, sheet.nrows):
    val = sheet.row_values(rowx)

    pos = PointofSales(
        code = val[0],
        description =  val[1],
        qty =  val[2],
        food_rev =  val[3],
        bev_rev =  val[4],
        oth_fb_rev =  val[5],
        amount =  val[6],
        sv_chg =  val[7],
        tax =  val[8],
        amount_net =  val[9],
        cash =  val[10],
        credit_card =  val[11],
        city_ledger =  val[12],
        bill_fo =  val[13],
        other_pymt =  val[14],
        remarks =  val[15],
        bill_no =  val[16]
        )
    session.add(pos)

session.commit()