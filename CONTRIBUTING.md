# Contributing to MS365-Access

Thank you for your interest in contributing to MS365-Access!

## Development Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/ms365-access.git
   cd ms365-access
   ```

2. **Create a virtual environment**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   cd backend
   pip install -r requirements.txt
   ```

4. **Configure environment**
   ```bash
   cp .env.example .env
   # Edit .env with your Azure AD credentials
   ```

5. **Run the development server**
   ```bash
   uvicorn app.main:app --port 8365 --reload
   ```

## Azure AD Setup

To develop locally, you need an Azure AD App Registration:

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Create a new registration
3. Add redirect URI: `http://localhost:8365/auth/callback`
4. Create a client secret
5. Add API permissions: User.Read, Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, Files.ReadWrite.All

## Running Tests

```bash
cd backend
pytest
```

## Code Style

- Follow PEP 8 guidelines
- Use type hints where possible
- Add docstrings to public functions

## Pull Request Process

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run tests to ensure nothing is broken
5. Commit your changes (`git commit -m "Add amazing feature"`)
6. Push to your branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

## Reporting Issues

When reporting issues, please include:
- Python version
- Operating system
- Steps to reproduce
- Expected vs actual behavior
- Relevant error messages or logs

## Security

If you discover a security vulnerability, please do NOT open a public issue. Instead, email the maintainer directly.

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
