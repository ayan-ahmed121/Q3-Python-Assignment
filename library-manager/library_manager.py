import pyodbc

class PersonalLibraryManager:
    def __init__(self, db_file):
        self.db_file = db_file
        self.conn = pyodbc.connect(f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={self.db_file};')
        self.cursor = self.conn.cursor()

    def add_book(self, title, author, year, genre, read_status):
        query = '''INSERT INTO Books (Title, Author, Year, Genre, ReadStatus)
                   VALUES (?, ?, ?, ?, ?)'''
        self.cursor.execute(query, (title, author, year, genre, read_status))
        self.conn.commit()
        print(f"Book '{title}' added successfully!")

    def remove_book(self, title):
        query = '''DELETE FROM Books WHERE Title = ?'''
        self.cursor.execute(query, (title,))
        self.conn.commit()
        print(f"Book '{title}' removed successfully!")

    def search_books(self, search_term):
        query = '''SELECT * FROM Books WHERE Title LIKE ? OR Author LIKE ? OR Genre LIKE ?'''
        self.cursor.execute(query, ('%' + search_term + '%', '%' + search_term + '%', '%' + search_term + '%'))
        books = self.cursor.fetchall()
        if books:
            for book in books:
                print(f"Title: {book.Title}, Author: {book.Author}, Year: {book.Year}, Genre: {book.Genre}, Read: {book.ReadStatus}")
        else:
            print("No books found matching the search term.")

    def view_books(self):
        query = '''SELECT * FROM Books'''
        self.cursor.execute(query)
        books = self.cursor.fetchall()
        if books:
            for book in books:
                print(f"Title: {book.Title}, Author: {book.Author}, Year: {book.Year}, Genre: {book.Genre}, Read: {book.ReadStatus}")
        else:
            print("No books in the library.")

    def display_statistics(self):
        query = '''SELECT COUNT(*) FROM Books'''
        self.cursor.execute(query)
        total_books = self.cursor.fetchone()[0]

        query = '''SELECT COUNT(*) FROM Books WHERE ReadStatus = True'''
        self.cursor.execute(query)
        read_books = self.cursor.fetchone()[0]

        if total_books > 0:
            read_percentage = (read_books / total_books) * 100
            print(f"Total Books: {total_books}")
            print(f"Books Read: {read_books} ({read_percentage:.2f}%)")
            print(f"Books Unread: {total_books - read_books}")
        else:
            print("No books to display statistics for.")

def main():
    library = PersonalLibraryManager(r'C:\Users\DELL\python\library-manager\Books.accdb')  

    while True:
        print("\nLibrary Menu:")
        print("1. Add a book")
        print("2. Remove a book")
        print("3. Search for a book")
        print("4. Display all books")
        print("5. Display statistics")
        print("6. Exit")

        choice = input("\nChoose an option (1-6): ")

        if choice == '1':
            title = input("Enter book title: ")
            author = input("Enter book author: ")
            year = int(input("Enter publication year: "))
            genre = input("Enter book genre: ")
            read_status = input("Is the book read (yes/no)? ").lower() == 'yes'
            library.add_book(title, author, year, genre, read_status)
        elif choice == '2':
            title = input("Enter book title to remove: ")
            library.remove_book(title)
        elif choice == '3':
            search_term = input("Enter search term (title/author/genre): ")
            library.search_books(search_term)
        elif choice == '4':
            print("\nLibrary Collection:")
            library.view_books()
        elif choice == '5':
            library.display_statistics()
        elif choice == '6':
            print("Exiting the Library Manager. Goodbye!")
            break
        else:
            print("Invalid choice, please try again.")

if __name__ == "__main__":
    main()
