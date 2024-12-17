# PostgreSQL to Excel Exporter

**pg-to-xlsx** is a utility for exporting data from a PostgreSQL database to an Excel (.xlsx) file. It allows you to execute arbitrary SQL queries and export the results directly to an Excel file.

---

## Features

- Connect to a PostgreSQL server.
- Execute any SQL query.
- Export query results to an Excel file.
- Support for various data types (strings, numbers, dates, etc.).

---

## Requirements

- [Go](https://golang.org/) version 1.17 or higher.
- The PostgreSQL driver for Go: `github.com/lib/pq`.
- Access to a PostgreSQL server with the necessary permissions.

---

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/dsalaidinov/pg-to-xlsx.git
   cd pg-to-xlsx
   ```

2. Build the program:

   ```bash
   go build -o pg-to-xlsx
   ```

3. After building, you will have the executable file `pg-to-xlsx`.

---

## Usage

Once the program is built, you can use it to export data from PostgreSQL to an Excel file.

The program accepts the following command-line parameters:

```bash
./pg-to-xlsx -h <host> -u <username> -p <password> -d <database_name> -s <sql_file> -o <output_file>
```

### Parameters:

- `-h <host>` — PostgreSQL server host (e.g., `localhost`).
- `-u <username>` — PostgreSQL username.
- `-p <password>` — PostgreSQL password.
- `-d <database_name>` — PostgreSQL database name.
- `-s <sql_file>` — Path to the SQL query file.
- `-o <output_file>` — Path to the output Excel file.

### Example:

```bash
./pg-to-xlsx -h localhost -u myuser -p mypassword -d mydb -s query.sql -o result.xlsx
```

This will execute the SQL query from the `query.sql` file, connect to the `mydb` database on the `localhost` server, and save the results to `result.xlsx`.

---

## Example SQL Query

The `query.sql` file can contain any valid SQL query, for example:

```sql
SELECT id, name, email, created_at FROM users WHERE active = TRUE;
```

The results of this query will be exported to Excel.

---

## Supported Data Types

The program supports the following data types, which will be correctly exported to Excel:

- Strings (`TEXT`, `VARCHAR`, `CHAR`).
- Numbers (`INTEGER`, `BIGINT`, `FLOAT`, `DECIMAL`).
- Boolean values (`BOOLEAN`).
- Dates and timestamps (`DATE`, `TIMESTAMP`).

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

---

## Contributing

If you have ideas for improvements or want to contribute to the code, please feel free to open a Pull Request. We are always happy to have new contributors!