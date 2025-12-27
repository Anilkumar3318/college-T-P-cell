"""
PDF Manager for Offer Letters
"""

import os
import tempfile
import subprocess
import platform
from datetime import datetime
from tkinter import messagebox
from pymongo import MongoClient
from bson import ObjectId

OFFER_LETTERS_URI = "mongodb+srv://Tp_offer_letter:TPoffer123@offerletters.axpsntb.mongodb.net/offer_letters"

class PDFManager:
    def __init__(self):
        self.client = None
        self.db = None
        self.collection = None
    
    def connect(self):
        try:
            if self.client is None:
                self.client = MongoClient(
                    OFFER_LETTERS_URI,
                    serverSelectionTimeoutMS=30000,
                    connectTimeoutMS=30000,
                    socketTimeoutMS=30000,
                    tls=True,
                    tlsAllowInvalidCertificates=True
                )
                self.db = self.client.offer_letters
                self.collection = self.db.letters
                
                self.client.admin.command('ping')
                print("✅ PDF Manager connected successfully!")
            return True
            
        except Exception as e:
            print(f"❌ PDF Manager connection failed: {e}")
            return False
    
    def store_pdf(self, file_path, student_name, company_name):
        """Store PDF in offer letters database"""
        try:
            if not self.connect():
                raise Exception("Failed to connect to database")
                
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"PDF file not found: {file_path}")
            
            # Check file size (max 100KB)
            file_size = os.path.getsize(file_path)
            if file_size > 102400:  # 100KB
                raise ValueError(f"File size {file_size/1024:.1f}KB exceeds 100KB limit")
            
            # Read PDF file as binary
            with open(file_path, 'rb') as pdf_file:
                pdf_data = pdf_file.read()
            
            # Create document
            document = {
                "student_name": student_name.upper(),
                "company_name": company_name.upper(),
                "filename": os.path.basename(file_path),
                "upload_date": datetime.now(),
                "pdf_data": pdf_data,
                "pdf_size": len(pdf_data),
                "status": "active"
            }
            
            # Check if PDF already exists for this student-company combination
            existing = self.collection.find_one({
                "student_name": student_name.upper(),
                "company_name": company_name.upper(),
                "status": "active"
            })
            
            if existing:
                # Update existing PDF
                self.collection.update_one(
                    {"_id": existing["_id"]},
                    {"$set": document}
                )
                pdf_id = str(existing["_id"])
                print(f"✅ Updated existing offer letter for {student_name} - {company_name}")
            else:
                # Insert new PDF
                result = self.collection.insert_one(document)
                pdf_id = str(result.inserted_id)
                print(f"✅ Stored new offer letter for {student_name} - {company_name}")
            
            return pdf_id
            
        except Exception as e:
            print(f"❌ Error storing PDF: {e}")
            messagebox.showerror("PDF Storage Error", f"Failed to store PDF: {e}")
            return None
    
    def _is_valid_objectid(self, pdf_id):
        """Check if string is a valid MongoDB ObjectId"""
        if not pdf_id or not isinstance(pdf_id, str):
            return False
        return len(pdf_id) == 24 and all(c in '0123456789abcdef' for c in pdf_id.lower())
    
    def pdf_exists(self, pdf_id):
        """Check if PDF exists in database by ID or by custom key"""
        try:
            if not pdf_id or not self.connect():
                return False
            
            # Try ObjectId first if valid format
            if self._is_valid_objectid(pdf_id):
                result = self.collection.find_one({
                    "_id": ObjectId(pdf_id),
                    "status": "active"
                })
                if result:
                    return True
            
            # Fallback: search by custom_key field (for old records)
            result = self.collection.find_one({
                "custom_key": pdf_id,
                "status": "active"
            })
            
            return result is not None
            
        except Exception as e:
            print(f"❌ Error checking PDF existence: {e}")
            return False
    
    def get_pdf_info(self, pdf_id):
        """Get PDF information without retrieving the actual file data"""
        try:
            if not pdf_id or not self.connect():
                return None
            
            result = None
            
            # Try ObjectId first if valid format
            if self._is_valid_objectid(pdf_id):
                result = self.collection.find_one(
                    {"_id": ObjectId(pdf_id), "status": "active"},
                    {"pdf_data": 0}
                )
            
            # Fallback: search by custom_key field
            if not result:
                result = self.collection.find_one(
                    {"custom_key": pdf_id, "status": "active"},
                    {"pdf_data": 0}
                )
            
            if result:
                return {
                    "filename": result.get("filename"),
                    "file_size": result.get("pdf_size"),
                    "upload_date": result.get("upload_date"),
                    "student_name": result.get("student_name"),
                    "company_name": result.get("company_name")
                }
            
            return None
            
        except Exception as e:
            print(f"❌ Error getting PDF info: {e}")
            return None
    
    def view_pdf(self, pdf_id):
        """View PDF by opening it in the default PDF viewer"""
        try:
            if not pdf_id or not self.connect():
                return {"success": False, "message": "Invalid PDF ID or connection failed"}
            
            result = None
            
            # Try ObjectId first if valid format
            if self._is_valid_objectid(pdf_id):
                result = self.collection.find_one({
                    "_id": ObjectId(pdf_id),
                    "status": "active"
                })
            
            # Fallback: search by custom_key field
            if not result:
                result = self.collection.find_one({
                    "custom_key": pdf_id,
                    "status": "active"
                })
            
            if not result:
                return {"success": False, "message": "PDF not found"}
            
            pdf_data = result.get("pdf_data")
            if not pdf_data:
                return {"success": False, "message": "PDF data not found"}
            
            # Create temporary file
            temp_file = tempfile.NamedTemporaryFile(
                delete=False, 
                suffix='.pdf',
                prefix=f"{result.get('student_name', 'offer')}_"
            )
            
            # Write PDF data to temporary file
            temp_file.write(pdf_data)
            temp_file.close()
            
            # Open PDF with default application
            try:
                if platform.system() == 'Windows':
                    os.startfile(temp_file.name)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', temp_file.name])
                else:  # Linux
                    subprocess.run(['xdg-open', temp_file.name])
                
                return {
                    "success": True, 
                    "message": f"Opened PDF for {result.get('student_name')} - {result.get('company_name')}",
                    "temp_file": temp_file.name
                }
                
            except Exception as open_error:
                # Clean up temp file if opening failed
                os.unlink(temp_file.name)
                return {"success": False, "message": f"Failed to open PDF: {open_error}"}
            
        except Exception as e:
            print(f"❌ Error viewing PDF: {e}")
            return {"success": False, "message": f"Error viewing PDF: {e}"}
    
    def delete_pdf(self, pdf_id):
        """Delete PDF from database (soft delete by setting status to inactive)"""
        try:
            if not pdf_id or not self.connect():
                return False
            
            # Try ObjectId first if valid format
            if self._is_valid_objectid(pdf_id):
                result = self.collection.update_one(
                    {"_id": ObjectId(pdf_id)},
                    {"$set": {"status": "inactive", "deleted_date": datetime.now()}}
                )
                if result.modified_count > 0:
                    return True
            
            # Fallback: search by custom_key field
            result = self.collection.update_one(
                {"custom_key": pdf_id},
                {"$set": {"status": "inactive", "deleted_date": datetime.now()}}
            )
            
            return result.modified_count > 0
            
        except Exception as e:
            print(f"❌ Error deleting PDF: {e}")
            return False
    
    def list_pdfs(self, student_name=None, company_name=None):
        """List PDFs with optional filtering"""
        try:
            if not self.connect():
                return []
            
            query = {"status": "active"}
            if student_name:
                query["student_name"] = student_name.upper()
            if company_name:
                query["company_name"] = company_name.upper()
            
            results = self.collection.find(
                query,
                {"pdf_data": 0}  # Exclude large PDF data
            ).sort("upload_date", -1)
            
            pdf_list = []
            for doc in results:
                pdf_info = {
                    "pdf_id": str(doc["_id"]),
                    "student_name": doc.get("student_name"),
                    "company_name": doc.get("company_name"),
                    "filename": doc.get("filename"),
                    "file_size": doc.get("pdf_size"),
                    "upload_date": doc.get("upload_date")
                }
                pdf_list.append(pdf_info)
            
            return pdf_list
            
        except Exception as e:
            print(f"❌ Error listing PDFs: {e}")
            return []
    
    def close(self):
        """Close database connection"""
        if self.client:
            self.client.close()
            self.client = None
            self.db = None
            self.collection = None

# Create global PDF manager instance
pdf_manager = PDFManager()

def test_pdf_manager():
    """Test the PDF manager functionality"""
    print("🧪 Testing Fixed PDF Manager")
    print("=" * 40)
    
    try:
        # Test connection
        if pdf_manager.connect():
            print("✅ Connection successful")
            
            # List existing PDFs
            pdfs = pdf_manager.list_pdfs()
            print(f"📋 Found {len(pdfs)} PDFs in database")
            
            for pdf in pdfs[:3]:  # Show first 3
                print(f"  - {pdf['student_name']} - {pdf['company_name']}")
                print(f"    ID: {pdf['pdf_id']}")
                print(f"    Size: {pdf['file_size']} bytes")
            
            return True
        else:
            print("❌ Connection failed")
            return False
            
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

if __name__ == "__main__":
    test_pdf_manager()