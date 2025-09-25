"""
Database Management
===================

SQLAlchemy-based database management with models for
cardholders, transactions, logs, and application state.
"""

from datetime import datetime
from typing import Optional, List, Any
from pathlib import Path

from sqlalchemy import create_engine, Column, Integer, String, DateTime, Float, Boolean, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship, Session
from sqlalchemy.pool import StaticPool

Base = declarative_base()

class Cardholder(Base):
    """Cardholder model"""
    __tablename__ = 'cardholders'
    
    id = Column(Integer, primary_key=True)
    card_number = Column(String(50), unique=True, nullable=False)
    name = Column(String(100), nullable=False)
    email = Column(String(100), nullable=False)
    manager_email = Column(String(100))
    department = Column(String(100))
    cost_centre = Column(String(50))
    active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationships
    transactions = relationship("Transaction", back_populates="cardholder")
    statements = relationship("Statement", back_populates="cardholder")

class Transaction(Base):
    """Transaction model"""
    __tablename__ = 'transactions'
    
    id = Column(Integer, primary_key=True)
    cardholder_id = Column(Integer, ForeignKey('cardholders.id'), nullable=False)
    transaction_date = Column(DateTime, nullable=False)
    merchant = Column(String(200))
    amount = Column(Float, nullable=False)
    currency = Column(String(10), default='GBP')
    category = Column(String(100))
    description = Column(Text)
    receipt_url = Column(String(500))
    approved = Column(Boolean, default=False)
    reconciled = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    
    # Relationships
    cardholder = relationship("Cardholder", back_populates="transactions")

class Statement(Base):
    """Statement model"""
    __tablename__ = 'statements'
    
    id = Column(Integer, primary_key=True)
    cardholder_id = Column(Integer, ForeignKey('cardholders.id'), nullable=False)
    period_start = Column(DateTime, nullable=False)
    period_end = Column(DateTime, nullable=False)
    total_amount = Column(Float, nullable=False)
    file_path = Column(String(500))
    sent_date = Column(DateTime)
    status = Column(String(50), default='pending')  # pending, sent, approved, rejected
    created_at = Column(DateTime, default=datetime.utcnow)
    
    # Relationships
    cardholder = relationship("Cardholder", back_populates="statements")

class EmailLog(Base):
    """Email log model"""
    __tablename__ = 'email_logs'
    
    id = Column(Integer, primary_key=True)
    recipient_email = Column(String(100), nullable=False)
    subject = Column(String(200), nullable=False)
    body_preview = Column(Text)
    sent_date = Column(DateTime, default=datetime.utcnow)
    status = Column(String(50), default='sent')  # sent, failed, bounced
    error_message = Column(Text)
    attachments_count = Column(Integer, default=0)

class ScriptExecution(Base):
    """Script execution log model"""
    __tablename__ = 'script_executions'
    
    id = Column(Integer, primary_key=True)
    script_name = Column(String(100), nullable=False)
    start_time = Column(DateTime, default=datetime.utcnow)
    end_time = Column(DateTime)
    status = Column(String(50), default='running')  # running, success, failed
    exit_code = Column(Integer)
    output = Column(Text)
    error_output = Column(Text)
    duration_seconds = Column(Float)

class UserSettings(Base):
    """User settings model"""
    __tablename__ = 'user_settings'
    
    id = Column(Integer, primary_key=True)
    setting_key = Column(String(100), unique=True, nullable=False)
    setting_value = Column(Text)
    data_type = Column(String(20), default='string')  # string, int, float, bool, json
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

class DatabaseManager:
    """Database management class"""
    
    def __init__(self, database_url: str = "sqlite:///data/dashboard.db"):
        self.database_url = database_url
        self.engine = None
        self.SessionLocal = None
        
        # Create data directory if using SQLite
        if database_url.startswith("sqlite"):
            db_path = database_url.replace("sqlite:///", "")
            Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        
        self._initialize()
    
    def _initialize(self):
        """Initialize database connection and tables"""
        # Create engine with connection pooling for SQLite
        if self.database_url.startswith("sqlite"):
            self.engine = create_engine(
                self.database_url,
                poolclass=StaticPool,
                connect_args={"check_same_thread": False, "timeout": 20}
            )
        else:
            self.engine = create_engine(self.database_url)
        
        # Create session factory
        self.SessionLocal = sessionmaker(
            autocommit=False, 
            autoflush=False, 
            bind=self.engine
        )
        
        # Create tables
        Base.metadata.create_all(bind=self.engine)
    
    def get_session(self) -> Session:
        """Get database session"""
        return self.SessionLocal()
    
    def close(self):
        """Close database connections"""
        if self.engine:
            self.engine.dispose()
    
    # Cardholder methods
    def create_cardholder(self, card_number: str, name: str, email: str,
                         manager_email: str = None, department: str = None,
                         cost_centre: str = None) -> Cardholder:
        """Create new cardholder"""
        with self.get_session() as session:
            cardholder = Cardholder(
                card_number=card_number,
                name=name,
                email=email,
                manager_email=manager_email,
                department=department,
                cost_centre=cost_centre
            )
            session.add(cardholder)
            session.commit()
            session.refresh(cardholder)
            return cardholder
    
    def get_cardholders(self, active_only: bool = True) -> List[Cardholder]:
        """Get all cardholders"""
        with self.get_session() as session:
            query = session.query(Cardholder)
            if active_only:
                query = query.filter(Cardholder.active == True)
            return query.all()
    
    def get_cardholder_by_card_number(self, card_number: str) -> Optional[Cardholder]:
        """Get cardholder by card number"""
        with self.get_session() as session:
            return session.query(Cardholder).filter(
                Cardholder.card_number == card_number
            ).first()
    
    # Transaction methods
    def create_transaction(self, cardholder_id: int, transaction_date: datetime,
                          merchant: str, amount: float, **kwargs) -> Transaction:
        """Create new transaction"""
        with self.get_session() as session:
            transaction = Transaction(
                cardholder_id=cardholder_id,
                transaction_date=transaction_date,
                merchant=merchant,
                amount=amount,
                **kwargs
            )
            session.add(transaction)
            session.commit()
            session.refresh(transaction)
            return transaction
    
    def get_transactions(self, cardholder_id: Optional[int] = None,
                        start_date: Optional[datetime] = None,
                        end_date: Optional[datetime] = None) -> List[Transaction]:
        """Get transactions with optional filters"""
        with self.get_session() as session:
            query = session.query(Transaction)
            
            if cardholder_id:
                query = query.filter(Transaction.cardholder_id == cardholder_id)
            if start_date:
                query = query.filter(Transaction.transaction_date >= start_date)
            if end_date:
                query = query.filter(Transaction.transaction_date <= end_date)
            
            return query.order_by(Transaction.transaction_date.desc()).all()
    
    # Email log methods
    def log_email(self, recipient_email: str, subject: str, body_preview: str = None,
                  status: str = 'sent', error_message: str = None,
                  attachments_count: int = 0) -> EmailLog:
        """Log email sending"""
        with self.get_session() as session:
            email_log = EmailLog(
                recipient_email=recipient_email,
                subject=subject,
                body_preview=body_preview,
                status=status,
                error_message=error_message,
                attachments_count=attachments_count
            )
            session.add(email_log)
            session.commit()
            session.refresh(email_log)
            return email_log
    
    # Script execution methods
    def log_script_start(self, script_name: str) -> ScriptExecution:
        """Log script execution start"""
        with self.get_session() as session:
            execution = ScriptExecution(
                script_name=script_name,
                status='running'
            )
            session.add(execution)
            session.commit()
            session.refresh(execution)
            return execution
    
    def log_script_end(self, execution_id: int, status: str, exit_code: int = None,
                       output: str = None, error_output: str = None):
        """Log script execution end"""
        with self.get_session() as session:
            execution = session.query(ScriptExecution).filter(
                ScriptExecution.id == execution_id
            ).first()
            
            if execution:
                execution.end_time = datetime.utcnow()
                execution.status = status
                execution.exit_code = exit_code
                execution.output = output
                execution.error_output = error_output
                
                if execution.start_time:
                    execution.duration_seconds = (
                        execution.end_time - execution.start_time
                    ).total_seconds()
                
                session.commit()
    
    # Settings methods
    def get_setting(self, key: str, default: Any = None) -> Any:
        """Get user setting"""
        with self.get_session() as session:
            setting = session.query(UserSettings).filter(
                UserSettings.setting_key == key
            ).first()
            
            if not setting:
                return default
            
            # Convert based on data type
            if setting.data_type == 'int':
                return int(setting.setting_value)
            elif setting.data_type == 'float':
                return float(setting.setting_value)
            elif setting.data_type == 'bool':
                return setting.setting_value.lower() in ['true', '1', 'yes']
            elif setting.data_type == 'json':
                import json
                return json.loads(setting.setting_value)
            else:
                return setting.setting_value
    
    def set_setting(self, key: str, value: Any):
        """Set user setting"""
        import json
        
        with self.get_session() as session:
            setting = session.query(UserSettings).filter(
                UserSettings.setting_key == key
            ).first()
            
            # Determine data type and convert value
            if isinstance(value, bool):
                data_type = 'bool'
                setting_value = str(value)
            elif isinstance(value, int):
                data_type = 'int'
                setting_value = str(value)
            elif isinstance(value, float):
                data_type = 'float'
                setting_value = str(value)
            elif isinstance(value, (dict, list)):
                data_type = 'json'
                setting_value = json.dumps(value)
            else:
                data_type = 'string'
                setting_value = str(value)
            
            if setting:
                setting.setting_value = setting_value
                setting.data_type = data_type
                setting.updated_at = datetime.utcnow()
            else:
                setting = UserSettings(
                    setting_key=key,
                    setting_value=setting_value,
                    data_type=data_type
                )
                session.add(setting)
            
            session.commit()

# Global database manager instance
db_manager = None

def get_db_manager(database_url: str = "sqlite:///data/dashboard.db") -> DatabaseManager:
    """Get global database manager instance"""
    global db_manager
    if db_manager is None:
        db_manager = DatabaseManager(database_url)
    return db_manager

__all__ = [
    'Base', 'Cardholder', 'Transaction', 'Statement', 'EmailLog',
    'ScriptExecution', 'UserSettings', 'DatabaseManager', 'get_db_manager'
]