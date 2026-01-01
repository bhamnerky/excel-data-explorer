#!/usr/bin/env python3
"""
Custom Query Runner for WIP Analysis
Interactive SQL query execution with save-to-CSV option
"""
import duckdb
import sys

conn = duckdb.connect('../wip_analysis.duckdb', read_only=True)


def show_schema():
    """Display table schema"""
    print("\n" + "="*80)
    print("📋 WIP Table Schema")
    print("="*80)
    schema = conn.execute("DESCRIBE wip").df()
    print(schema.to_string(index=False))
    print(f"\n✓ {len(schema)} columns total")


def show_sample_queries():
    """Show example queries"""
    print("\n" + "="*80)
    print("💡 Example Queries")
    print("="*80)
    examples = [
        ("Top contracts by revenue",
         'SELECT Contract, "Customer Name", "Revenue To Date" FROM wip ORDER BY "Revenue To Date" DESC LIMIT 10'),
        ("Contracts by region and status",
         'SELECT Region, "Contract Status", COUNT(*) FROM wip GROUP BY Region, "Contract Status"'),
        ("Average margin by service type",
         'SELECT ServiceType, AVG("Gross Profit %") * 100 as avg_margin FROM wip GROUP BY ServiceType'),
        ("Open contracts with completion >90%",
         'SELECT Contract, "% Complete" FROM wip WHERE "Contract Status" = \'Open\' AND "% Complete" > 0.9')
    ]
    for i, (desc, query) in enumerate(examples, 1):
        print(f"\n{i}. {desc}:")
        print(f"   {query}")


def run_custom_query(query):
    """Execute custom SQL query"""
    try:
        result = conn.execute(query).df()
        print("\n" + "="*80)
        print("📊 Query Results")
        print("="*80)
        print(result.to_string(index=False))
        print(f"\n✓ {len(result)} rows returned")

        # Offer to save
        if len(result) > 0:
            save = input("\n💾 Save to CSV? (y/n): ").strip().lower()
            if save == 'y':
                filename = input("Filename (without .csv): ").strip()
                if filename:
                    csv_path = f"{filename}.csv"
                    result.to_csv(csv_path, index=False)
                    print(f"✓ Saved to {csv_path}")
    except Exception as e:
        print(f"❌ Error: {e}")
        print("\nTip: Use double quotes for column names with spaces")
        print('Example: SELECT "Revenue To Date" FROM wip')


def interactive_mode():
    """Interactive query mode"""
    print("\n" + "="*80)
    print("🦆 DuckDB Interactive Query Mode")
    print("="*80)
    print("\nCommands:")
    print("  schema  - Show table structure")
    print("  examples - Show sample queries")
    print("  quit    - Exit")
    print("\nOr enter any SQL query to execute")

    while True:
        try:
            query = input("\nSQL> ").strip()

            if query.lower() in ['quit', 'exit', 'q']:
                break
            elif query.lower() == 'schema':
                show_schema()
            elif query.lower() in ['examples', 'help']:
                show_sample_queries()
            elif query:
                run_custom_query(query)
        except KeyboardInterrupt:
            print("\n\n✓ Exiting...")
            break
        except EOFError:
            break


if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Run query from command line argument
        query = " ".join(sys.argv[1:])
        run_custom_query(query)
    else:
        # Interactive mode
        interactive_mode()

conn.close()
