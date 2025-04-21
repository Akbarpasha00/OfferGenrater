# OfferGenrater

A Flask-based service to generate personalized offer letters.

## Files
- `letter_generator.py` — main application
- `requirements.txt` — dependencies
- `Procfile` — for Railway deployment

## Usage
1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Run Flask app (development):
   ```
   python letter_generator.py
   ```
3. Deploy to Railway by pushing to a connected GitHub repo.  
   Railway will use the Procfile to start the app.