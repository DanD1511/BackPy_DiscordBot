# BackPy Discord Bot

## Description
BackPy is a Discord bot designed to automate the generation of documents based on user input and attachments within Discord. It streamlines processes such as requesting specific panel types (640, 600, 580) and generating documents based on these specifications and provided calculation memories.

## Features
- Interactive document generation based on panel type selection.
- Supports various templates for document generation.
- Handles user inputs and file attachments to produce custom documents.
- Provides feedback and prompts to guide the user through the document generation process.

## Installation

### Prerequisites
- Python 3.6+
- discord.py
- aiohttp
- openpyxl
- python-dotenv
- DocxTemplate

### Setup
1. Clone the repository to your local machine.
2. Install the required dependencies by running `pip install -r requirements.txt`.
3. Create a `.env` file in the root directory and add your Discord bot token as `DISCORD_TOKEN=your_token_here`.
4. Run the bot with `python main.py`.

## Usage
To use the bot, type the `!OR` command in any text channel the bot has access to. Follow the interactive prompts to select a panel type and upload a calculation memory. The bot will then generate and send the document.

## Contributing
Contributions are welcome! If you have a bug fix or a new feature you'd like to add, please open a pull request.

## License
Please add your license information here.

For more detailed usage instructions and information on how to contribute to the project, please check the [GitHub repository](https://github.com/DanD1511/BackPy_DiscordBot).
