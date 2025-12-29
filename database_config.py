import pymongo
from pymongo import MongoClient
from tkinter import messagebox
import threading
from functools import lru_cache

MONGO_URI = "DATABASE URI"
DATABASE_NAME = "TPinfo"

OFFER_LETTERS_URI = "DATABASE URI"
OFFER_LETTERS_DB_NAME = "offer_letters"

COLLECTIONS = {
    "student": "student",
    "company": "company", 
    "placed_student": "placed_student"
}

OFFER_LETTERS_COLLECTION = "letters"

class OptimizedDatabaseManager:
    _instance = None
    _lock = threading.Lock()
    
    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, 'initialized'):
            self.client = None
            self.db = None
            self.offer_letters_client = None
            self.offer_letters_db = None
            self._collections = {}
            self.initialized = True
    
    def connect(self):
        if self.client is None:
            try:
                self.client = MongoClient(
                    MONGO_URI,
                    maxPoolSize=15,
                    minPoolSize=2,
                    maxIdleTimeMS=45000,
                    serverSelectionTimeoutMS=20000,
                    connectTimeoutMS=20000,
                    socketTimeoutMS=20000,
                    tls=True,
                    tlsAllowInvalidCertificates=True,
                    retryWrites=True,
                    w='majority'
                )
                
                self.client.admin.command('ping')
                self.db = self.client[DATABASE_NAME]
                print("✅ Main DB connection successful!")
                return True
                
            except Exception as e:
                error_msg = f"Failed to connect to main MongoDB: {e}"
                print(f"❌ {error_msg}")
                messagebox.showerror("Database Error", error_msg)
                return False
        return True
    
    def connect_offer_letters(self):
        """Establish connection to offer letters database"""
        if self.offer_letters_client is None:
            try:
                self.offer_letters_client = MongoClient(
                    OFFER_LETTERS_URI,
                    maxPoolSize=8,   # Smaller pool for offer letters
                    minPoolSize=1,
                    maxIdleTimeMS=45000,
                    serverSelectionTimeoutMS=20000,
                    connectTimeoutMS=20000,
                    socketTimeoutMS=20000,
                    tls=True,
                    tlsAllowInvalidCertificates=True,
                    retryWrites=True
                )
                
                # Test the connection
                self.offer_letters_client.admin.command('ping')
                self.offer_letters_db = self.offer_letters_client[OFFER_LETTERS_DB_NAME]
                print("✅ Offer letters DB connection successful!")
                return True
                
            except Exception as e:
                error_msg = f"Failed to connect to offer letters MongoDB: {e}"
                print(f"❌ {error_msg}")
                messagebox.showerror("Offer Letters Database Error", error_msg)
                return False
        return True
    
    @lru_cache(maxsize=3)  # Cache collections for faster access
    def get_collection(self, collection_name):
        """Get collection with caching"""
        if collection_name not in COLLECTIONS:
            raise ValueError(f"Unknown collection: {collection_name}")
            
        if not self.connect():
            return None
            
        # Return cached collection 
        if collection_name not in self._collections:
            self._collections[collection_name] = self.db[COLLECTIONS[collection_name]]
            
        return self._collections[collection_name]
    
    def get_offer_letters_collection(self):
        """Get offer letters collection from separate database"""
        if not self.connect_offer_letters():
            return None
        return self.offer_letters_db[OFFER_LETTERS_COLLECTION]
    
    def close_connection(self):
        """Close database connection and clear cache"""
        if self.client:
            self.client.close()
            self.client = None
            self.db = None
        
        if self.offer_letters_client:
            self.offer_letters_client.close()
            self.offer_letters_client = None
            self.offer_letters_db = None
            
        self._collections.clear()
        self.get_collection.cache_clear()



db_manager = OptimizedDatabaseManager()


def get_student_collection():
    """Get student collection"""
    return db_manager.get_collection("student")


def get_company_collection():
    """Get company collection"""
    return db_manager.get_collection("company")


def get_placed_student_collection():
    """Get placed student collection"""
    return db_manager.get_collection("placed_student")


def get_offer_letters_collection():
    """Get offer letters collection from separate database"""
    return db_manager.get_offer_letters_collection()


def test_connection():
    """Test database connection"""
    try:
        # Test main database
        if db_manager.connect():
            collections = db_manager.db.list_collection_names()
            print(f"✅ Main database connection successful!")
            print(f"📊 Available collections: {collections}")
            
            # Test offer letters database
            if db_manager.connect_offer_letters():
                offer_collections = db_manager.offer_letters_db.list_collection_names()
                print(f"✅ Offer letters database connection successful!")
                print(f"📄 Offer letters collections: {offer_collections}")
                return True
            else:
                print("❌ Offer letters database connection failed!")
                return False
        else:
            print("❌ Main database connection failed!")
            return False
    except Exception as e:
        print(f"❌ Database connection test failed: {e}")
        return False


if __name__ == "__main__":
    test_connection()

