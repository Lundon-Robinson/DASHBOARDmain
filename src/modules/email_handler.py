"""
Email Handler Module
====================

Bulk email functionality with templates, tracking, and
Outlook/Teams integration.
"""

import re
import smtplib
from pathlib import Path
from typing import List, Dict, Any, Optional
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import uuid

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

from ..core.logger import logger
from ..core.database import get_db_manager
from ..core.config import DashboardConfig

class EmailTemplate:
    """Email template class"""
    
    def __init__(self, name: str, subject: str, body: str, 
                 variables: List[str] = None):
        self.name = name
        self.subject = subject
        self.body = body
        self.variables = variables or []
        self.created_at = datetime.now()

class EmailRecipient:
    """Email recipient information"""
    
    def __init__(self, email: str, name: str = "", variables: Dict[str, str] = None):
        self.email = email
        self.name = name
        self.variables = variables or {}

class EmailJob:
    """Email sending job"""
    
    def __init__(self, template: EmailTemplate, recipients: List[EmailRecipient],
                 attachments: List[str] = None, priority: str = "normal"):
        self.id = str(uuid.uuid4())
        self.template = template
        self.recipients = recipients
        self.attachments = attachments or []
        self.priority = priority
        self.created_at = datetime.now()
        self.started_at: Optional[datetime] = None
        self.completed_at: Optional[datetime] = None
        self.status = "pending"  # pending, running, completed, failed
        self.results: List[Dict[str, Any]] = []

class EmailHandler:
    """Advanced email processing handler"""
    
    def __init__(self, config: DashboardConfig):
        self.config = config
        self.db_manager = get_db_manager(config.database.url)
        
        # Email templates storage
        self.templates: Dict[str, EmailTemplate] = {}
        self.template_dir = config.paths.templates_dir / "email"
        self.template_dir.mkdir(parents=True, exist_ok=True)
        
        # Job queue
        self.job_queue: List[EmailJob] = []
        
        # Load default templates
        self._load_default_templates()
        
        logger.info("Email handler initialized")
        if WIN32_AVAILABLE:
            logger.info("Outlook integration available")
        else:
            logger.warning("Outlook integration not available (win32com not found)")
    
    def _load_default_templates(self):
        """Load default email templates"""
        # Purchase card statement template
        statement_template = EmailTemplate(
            name="purchase_card_statement",
            subject="Purchase Card Statement - {period}",
            body="""Dear {cardholder_name},

Please find attached your purchase card statement for the period {period}.

Card Number: {card_number}
Total Amount: {total_amount}

Please review the statement and submit any required documentation.

If you have any questions, please contact the Finance team.

Best regards,
Finance Team""",
            variables=["cardholder_name", "period", "card_number", "total_amount"]
        )
        self.templates["purchase_card_statement"] = statement_template
        
        # Delegation update template
        delegation_template = EmailTemplate(
            name="delegation_update",
            subject="Delegation Update Required - {delegated_officer}",
            body="""Dear {manager_name},

A delegation update is required for {delegated_officer}.

Please review and update the delegation limits and approvals as necessary.

Officer: {delegated_officer}
Department: {department}
Current Limits: {current_limits}

Please process this update at your earliest convenience.

Best regards,
Administration Team""",
            variables=["manager_name", "delegated_officer", "department", "current_limits"]
        )
        self.templates["delegation_update"] = delegation_template
        
        logger.debug(f"Loaded {len(self.templates)} default templates")
    
    def create_template(self, name: str, subject: str, body: str, 
                       variables: List[str] = None) -> EmailTemplate:
        """Create a new email template"""
        template = EmailTemplate(name, subject, body, variables)
        self.templates[name] = template
        
        # Save to file
        self._save_template(template)
        
        logger.info(f"Created email template: {name}")
        return template
    
    def _save_template(self, template: EmailTemplate):
        """Save template to file"""
        template_path = self.template_dir / f"{template.name}.json"
        
        import json
        template_data = {
            'name': template.name,
            'subject': template.subject,
            'body': template.body,
            'variables': template.variables,
            'created_at': template.created_at.isoformat()
        }
        
        with open(template_path, 'w', encoding='utf-8') as f:
            json.dump(template_data, f, indent=2)
    
    def load_template(self, name: str) -> Optional[EmailTemplate]:
        """Load template by name"""
        if name in self.templates:
            return self.templates[name]
        
        # Try to load from file
        template_path = self.template_dir / f"{name}.json"
        if template_path.exists():
            import json
            with open(template_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            template = EmailTemplate(
                data['name'],
                data['subject'],
                data['body'],
                data.get('variables', [])
            )
            self.templates[name] = template
            return template
        
        return None
    
    def list_templates(self) -> List[str]:
        """Get list of available templates"""
        return list(self.templates.keys())
    
    def create_bulk_job(self, template_name: str, recipients: List[EmailRecipient],
                       attachments: List[str] = None, priority: str = "normal") -> str:
        """Create a bulk email job"""
        template = self.load_template(template_name)
        if not template:
            raise ValueError(f"Template '{template_name}' not found")
        
        job = EmailJob(template, recipients, attachments, priority)
        self.job_queue.append(job)
        
        logger.info(f"Created bulk email job {job.id} with {len(recipients)} recipients")
        return job.id
    
    def process_variables(self, text: str, variables: Dict[str, str]) -> str:
        """Process template variables in text"""
        processed = text
        
        # Replace {variable} patterns
        for var_name, var_value in variables.items():
            pattern = f"{{{var_name}}}"
            processed = processed.replace(pattern, str(var_value))
        
        # Add common variables
        common_vars = {
            "today": datetime.now().strftime("%d/%m/%Y"),
            "full_date": datetime.now().strftime("%A, %B %d, %Y"),
            "time": datetime.now().strftime("%H:%M:%S"),
            "organisation": "Finance Department"
        }
        
        for var_name, var_value in common_vars.items():
            pattern = f"{{{var_name}}}"
            processed = processed.replace(pattern, str(var_value))
        
        return processed
    
    def send_via_outlook(self, to_email: str, subject: str, body: str,
                        attachments: List[str] = None, cc_email: str = "",
                        urgent: bool = False) -> bool:
        """Send email via Outlook"""
        if not WIN32_AVAILABLE:
            logger.error("Outlook not available - cannot send email")
            return False
        
        try:
            outlook = win32.Dispatch("Outlook.Application")
            
            # Create email
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.To = to_email
            
            if cc_email:
                mail.CC = cc_email
            
            mail.Subject = subject
            mail.Body = body
            
            # Set priority
            if urgent:
                mail.Importance = 2  # High importance
            
            # Add attachments
            if attachments:
                for attachment_path in attachments:
                    if Path(attachment_path).exists():
                        mail.Attachments.Add(str(attachment_path))
                    else:
                        logger.warning(f"Attachment not found: {attachment_path}")
            
            # Send email
            mail.Send()
            
            # Log to database
            self.db_manager.log_email(
                recipient_email=to_email,
                subject=subject,
                body_preview=body[:200] + "..." if len(body) > 200 else body,
                status="sent",
                attachments_count=len(attachments) if attachments else 0
            )
            
            logger.info(f"Email sent via Outlook to {to_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send email via Outlook to {to_email}", exception=e)
            
            # Log failed attempt
            self.db_manager.log_email(
                recipient_email=to_email,
                subject=subject,
                body_preview=body[:200] + "..." if len(body) > 200 else body,
                status="failed",
                error_message=str(e),
                attachments_count=len(attachments) if attachments else 0
            )
            
            return False
    
    def send_via_smtp(self, to_email: str, subject: str, body: str,
                     attachments: List[str] = None, is_html: bool = False) -> bool:
        """Send email via SMTP"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = self.config.email.username
            msg['To'] = to_email
            msg['Subject'] = subject
            
            # Add body
            if is_html:
                msg.attach(MIMEText(body, 'html'))
            else:
                msg.attach(MIMEText(body, 'plain'))
            
            # Add attachments
            if attachments:
                for attachment_path in attachments:
                    if Path(attachment_path).exists():
                        with open(attachment_path, "rb") as attachment:
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(attachment.read())
                        
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename= {Path(attachment_path).name}'
                        )
                        msg.attach(part)
                    else:
                        logger.warning(f"Attachment not found: {attachment_path}")
            
            # Connect to server and send
            server = smtplib.SMTP(self.config.email.smtp_server, self.config.email.smtp_port)
            
            if self.config.email.use_tls:
                server.starttls()
            
            if self.config.email.username and self.config.email.password:
                server.login(self.config.email.username, self.config.email.password)
            
            server.send_message(msg)
            server.quit()
            
            # Log success
            self.db_manager.log_email(
                recipient_email=to_email,
                subject=subject,
                body_preview=body[:200] + "..." if len(body) > 200 else body,
                status="sent",
                attachments_count=len(attachments) if attachments else 0
            )
            
            logger.info(f"Email sent via SMTP to {to_email}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send email via SMTP to {to_email}", exception=e)
            
            # Log failure
            self.db_manager.log_email(
                recipient_email=to_email,
                subject=subject,
                body_preview=body[:200] + "..." if len(body) > 200 else body,
                status="failed",
                error_message=str(e),
                attachments_count=len(attachments) if attachments else 0
            )
            
            return False
    
    def execute_bulk_job(self, job_id: str, use_outlook: bool = True,
                        delay_seconds: float = 1.0) -> Dict[str, Any]:
        """Execute a bulk email job"""
        job = None
        for j in self.job_queue:
            if j.id == job_id:
                job = j
                break
        
        if not job:
            raise ValueError(f"Job {job_id} not found")
        
        job.started_at = datetime.now()
        job.status = "running"
        
        success_count = 0
        failed_count = 0
        
        logger.info(f"Starting bulk email job {job_id} with {len(job.recipients)} recipients")
        
        for i, recipient in enumerate(job.recipients):
            try:
                # Process template
                subject = self.process_variables(job.template.subject, recipient.variables)
                body = self.process_variables(job.template.body, recipient.variables)
                
                # Send email
                if use_outlook and WIN32_AVAILABLE:
                    success = self.send_via_outlook(
                        recipient.email, subject, body, job.attachments
                    )
                else:
                    success = self.send_via_smtp(
                        recipient.email, subject, body, job.attachments
                    )
                
                # Record result
                job.results.append({
                    'recipient': recipient.email,
                    'success': success,
                    'timestamp': datetime.now()
                })
                
                if success:
                    success_count += 1
                else:
                    failed_count += 1
                
                # Add delay between emails
                if delay_seconds > 0 and i < len(job.recipients) - 1:
                    import time
                    time.sleep(delay_seconds)
                
            except Exception as e:
                logger.error(f"Failed to send email to {recipient.email}", exception=e)
                failed_count += 1
                
                job.results.append({
                    'recipient': recipient.email,
                    'success': False,
                    'error': str(e),
                    'timestamp': datetime.now()
                })
        
        job.completed_at = datetime.now()
        job.status = "completed" if failed_count == 0 else "partial"
        
        results = {
            'job_id': job_id,
            'total_recipients': len(job.recipients),
            'success_count': success_count,
            'failed_count': failed_count,
            'duration_seconds': (job.completed_at - job.started_at).total_seconds(),
            'results': job.results
        }
        
        logger.info(f"Bulk email job {job_id} completed: {success_count} sent, {failed_count} failed")
        return results
    
    def get_job_status(self, job_id: str) -> Dict[str, Any]:
        """Get status of email job"""
        job = None
        for j in self.job_queue:
            if j.id == job_id:
                job = j
                break
        
        if not job:
            raise ValueError(f"Job {job_id} not found")
        
        return {
            'job_id': job.id,
            'status': job.status,
            'created_at': job.created_at,
            'started_at': job.started_at,
            'completed_at': job.completed_at,
            'recipient_count': len(job.recipients),
            'results_count': len(job.results)
        }
    
    def find_attachments_by_name(self, folder_path: str, 
                                name_pattern: str) -> List[str]:
        """Find attachment files by name pattern"""
        folder = Path(folder_path)
        if not folder.exists():
            return []
        
        # Convert pattern to regex
        regex_pattern = name_pattern.replace('*', '.*').replace('?', '.')
        pattern = re.compile(regex_pattern, re.IGNORECASE)
        
        attachments = []
        for file_path in folder.rglob('*'):
            if file_path.is_file() and pattern.match(file_path.name):
                attachments.append(str(file_path))
        
        logger.debug(f"Found {len(attachments)} files matching pattern '{name_pattern}'")
        return attachments
    
    def get_email_statistics(self) -> Dict[str, Any]:
        """Get email sending statistics"""
        try:
            with self.db_manager.get_session() as session:
                from ..core.database import EmailLog
                
                # Total emails
                total_sent = session.query(EmailLog).filter(
                    EmailLog.status == 'sent'
                ).count()
                
                total_failed = session.query(EmailLog).filter(
                    EmailLog.status == 'failed'
                ).count()
                
                # Recent activity (last 7 days)
                week_ago = datetime.now() - timedelta(days=7)
                recent_sent = session.query(EmailLog).filter(
                    EmailLog.sent_date >= week_ago,
                    EmailLog.status == 'sent'
                ).count()
                
                return {
                    'total_sent': total_sent,
                    'total_failed': total_failed,
                    'recent_sent': recent_sent,
                    'success_rate': total_sent / (total_sent + total_failed) * 100 if (total_sent + total_failed) > 0 else 0
                }
        except Exception as e:
            logger.error("Failed to get email statistics", exception=e)
            return {
                'total_sent': 0,
                'total_failed': 0,
                'recent_sent': 0,
                'success_rate': 0
            }
    
    def validate_email(self, email: str) -> bool:
        """Validate email address format"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def parse_recipient_list(self, recipients_text: str) -> List[EmailRecipient]:
        """Parse recipient list from text (comma/semicolon separated)"""
        recipients = []
        
        # Split by common separators
        email_list = re.split(r'[,;\n\r]+', recipients_text)
        
        for email_entry in email_list:
            email_entry = email_entry.strip()
            if not email_entry:
                continue
            
            # Check for "Name <email>" format
            match = re.match(r'^(.+?)\s*<(.+?)>$', email_entry)
            if match:
                name = match.group(1).strip()
                email = match.group(2).strip()
            else:
                name = ""
                email = email_entry
            
            if self.validate_email(email):
                recipients.append(EmailRecipient(email, name))
            else:
                logger.warning(f"Invalid email address: {email}")
        
        return recipients

__all__ = ['EmailHandler', 'EmailTemplate', 'EmailRecipient', 'EmailJob']