from sqlalchemy import Column,Integer,ForeignKey,DateTime,String
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()

class Project(Base):
	__tablename__ = 'project'
	id = Column(Integer, primary_key=True)
	name = Column(String(350), nullable=False)

class Designation(Base):
	__tablename__ = 'designation'
	id = Column(Integer, primary_key=True)
	name = Column(String(350), nullable=False)

class Transaction(Base):
	__tablename__ = 'transaction'
	id = Column(Integer, primary_key=True)
	tr_date = Column(DateTime)
	project_id = Column(Integer, ForeignKey('project.id'))
	project = relationship(Project)
	remarks = Column(String(350), nullable=True)
	
class TransactionDetails(Base):
	__tablename__ = 'transactiondetails'
	id = Column(Integer, primary_key=True)
	transaction_id = Column(Integer, ForeignKey('transaction.id'))
	transaction = relationship(Transaction)
	designation_id = Column(Integer, ForeignKey('designation.id'))
	designation = relationship(Designation)
	present = Column(Integer)
	absent = Column(Integer)
	vacation = Column(Integer)
