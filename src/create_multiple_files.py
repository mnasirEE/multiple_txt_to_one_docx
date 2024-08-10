import os

def create_sample_files(num_files, directory='sample_files'):
    # Ensure the output directory exists
    if not os.path.exists(directory):
        os.makedirs(directory)
    
    for i in range(1, num_files + 1):
        filename = f'sample_file_{i}.txt'
        filepath = os.path.join(directory, filename)
        
        # Create and write content to the file
        with open(filepath, 'w') as file:
            file.write(f"Sample configuration content for file {i}.\n")
            file.write(f"hostname SampleHost_{i}\n")
            file.write("Additional content...\n")
            file.write("no ip http server\n")
            file.write("More configuration...\n")
    
    print(f"{num_files} sample files created in the '{directory}' directory.")

# Create 100 sample files
create_sample_files(100)
