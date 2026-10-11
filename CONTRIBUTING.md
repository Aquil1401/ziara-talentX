# Contributing to Ziara TalentX

First off, thank you for considering contributing to **Ziara TalentX**! It is contributions from developers like you that make recruitment automation more reliable, privacy-preserving, and accessible.

---

## 📋 Code of Conduct
By participating in this project, you agree to abide by our respectful, professional community guidelines. Please treat all contributors with respect and professionalism.

---

## 🛠️ How Can I Contribute?

### 1. Reporting Bugs
- Check the [existing issues](https://github.com/Aquil1401/ziara-talentX/issues) to avoid duplicate reports.
- Use our structured **Bug Report Template** to submit detailed information including OS, Python version, sample file type, and error logs.
- **CRITICAL:** Never attach real candidate resumes or personal details. Always reproduce issues using synthetic dummy files!

### 2. Suggesting Features
- We welcome feature suggestions (e.g., new file format parsers, cloud ATS exporters, AI scoring engines).
- Submit suggestions via our **Feature Request Template**.

### 3. Code Contributions (Pull Requests)
We follow standard GitHub Fork & Pull Request workflows:

1. **Fork the Repository:** Create your own fork of `Aquil1401/ziara-talentX`.
2. **Clone Locally:**
   ```bash
   git clone https://github.com/YOUR_USERNAME/ziara-talentX.git
   cd ziara-talentX
   ```
3. **Create a Virtual Environment:**
   ```bash
   python -m venv venv
   # Windows:
   venv\Scripts\activate
   # Linux/macOS:
   source venv/bin/activate
   ```
4. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
5. **Create a Feature Branch:**
   ```bash
   git checkout -b feat/your-feature-name
   ```
6. **Make Changes & Test:**
   - Run compilation check: `python -m py_compile extract_resumes.py`
   - Test on sample candidates: `python extract_resumes.py --folder candidate_sample`
   - Ensure CSV and Excel files are generated cleanly.
7. **Commit Your Code:**
   Follow Conventional Commits guidelines:
   - `feat:` for new capabilities
   - `fix:` for bug fixes
   - `docs:` for documentation updates
   - `refactor:` for code cleanups without behavior changes
8. **Open a Pull Request:**
   - Push your branch to GitHub and open a PR against `main`.
   - Fill out the PR template completely.

---

## ⚖️ License & Intellectual Property
By contributing code to Ziara TalentX, you agree that your contributions will be licensed under the project's [PolyForm Noncommercial License 1.0.0](LICENSE).
