from transformers import pipeline

print("Downloading model... this might take a minute...")
# We explicitly allow internet here to refresh the cache
pipeline("zero-shot-classification", model="valhalla/distilbart-mnli-12-1")
print("Download complete! You can now run your main script.")