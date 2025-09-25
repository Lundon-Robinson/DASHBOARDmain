"""
Email Handler Module
====================

Bulk email functionality with templates, tracking, and
Outlook/Teams integration.
"""

class EmailHandler:
    """Email processing handler"""
    
    def __init__(self, config):
        self.config = config