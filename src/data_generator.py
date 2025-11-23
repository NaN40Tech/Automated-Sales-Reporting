import pandas as pd
import random
from datetime import datetime, timedelta
import os

def generate_data(output_path):
    """Generates dummy sales data to the specified path."""
    
    products = {
        'Laptop': {'price': 12000000, 'category': 'Electronics'},
        'Mouse': {'price': 150000, 'category': 'Accessories'},
        'Keyboard': {'price': 350000, 'category': 'Accessories'},
        'Monitor': {'price': 2500000, 'category': 'Electronics'},
        'Chair': {'price': 1500000, 'category': 'Furniture'},
        'Desk': {'price': 3000000, 'category': 'Furniture'},
        'Headset': {'price': 500000, 'category': 'Accessories'}
    }
    
    regions = ['Jakarta', 'Bandung', 'Surabaya', 'Medan', 'Bali']
    
    data = []
    start_date = datetime(2024, 1, 1)
    
    for _ in range(100):
        date = start_date + timedelta(days=random.randint(0, 30))
        product_name = random.choice(list(products.keys()))
        product_info = products[product_name]
        
        row = {
            'Date': date.strftime('%Y-%m-%d'),
            'Product': product_name,
            'Category': product_info['category'],
            'Quantity': random.randint(1, 5),
            'UnitPrice': product_info['price'],
            'Region': random.choice(regions)
        }
        data.append(row)
        
    df = pd.DataFrame(data)
    
    # Ensure directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    df.to_csv(output_path, index=False)
    print(f"Data generated successfully at {output_path}")

if __name__ == "__main__":
    # Default behavior for testing
    generate_data('raw_sales_data.csv')
