# How to Run the Python Script to Create a Presentation

This guide will help you run the Python script to create the PPT deck code in `create_ppt_deck.py`

## Prerequisites

Before running the script, ensure you have the following:

1. Install Git from Git [docs here](https://git-scm.com/downloads)

2. **Install Python** & **Enable Python Virtual Environment**:

- Make sure Python (version 3.8 or higher) is installed on your computer. You can download it from [python.org](https://www.python.org/downloads/).
- It's recommended to use a virtual environment to manage dependencies. This guide assumes you are using a virtual environment.

## Steps to Run the Script

### 1. Clone the Repository

First, clone the repository containing the `create_ppt_deck.py` script to your local machine. Open your terminal or command prompt and run the following command:

```sh
git clone https://github.com/username/repo_name.git
```

Replace `username` and `repo_name` with the actual GitHub username and repository name.

### 2. Navigate to the Cloned Directory

Change to the directory of the cloned repository:

```sh
cd repo_name
```

### 3. Set Up the Virtual Environment

Set up a virtual environment:

```sh
python -m venv venv
```

Activate the virtual environment:

- **Windows**:
  ```sh
  venv\Scripts\activate
  ```
- **macOS/Linux**:
  ```sh
  source venv/bin/activate
  ```

### 4. Install Required Dependencies

Install the required Python packages listed in the 

requirements.txt

 file:

```sh
pip install -r requirements.txt
```

### 5. Run the Script

Run the script to create the PowerPoint presentation:

```sh
python create_ppt_deck.py
```

### 6. Verify the Output

After running the script, a file named `My_Deck.pptx` will be created in the same directory. Open this file using Microsoft PowerPoint or any compatible presentation software to verify the content.

## Troubleshooting

- **Python Not Recognized**: If you receive an error stating that Python is not recognized, ensure Python is added to your system's PATH.
- **Dependency Issues**: If there are issues installing dependencies, ensure you are using the correct version of Python and that your virtual environment is activated.

## Additional Information

- **Script Overview**: The script creates a PowerPoint presentation with multiple slides, each styled with specific fonts, colors, and background settings.
- **Customization**: You can customize the content and styling by modifying the `slides_content` list and the `styling` functions in the script.

By following these steps, you should be able to successfully run the script and generate a PowerPoint presentation. If you encounter any issues, refer to the troubleshooting section or seek assistance from a technical expert.
