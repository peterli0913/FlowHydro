from faster_whisper import WhisperModel

model = WhisperModel("small", device="cpu", compute_type="int8")
segments, info = model.transcribe(
    "/tmp/mtg0723.wav",
    language="zh",
    vad_filter=True,
    vad_parameters=dict(min_silence_duration_ms=500),
    beam_size=5,
)
print("detected lang:", info.language, "prob:", round(info.language_probability, 2))
with open("/tmp/transcript0723.txt", "w", encoding="utf-8") as f:
    for seg in segments:
        line = f"[{int(seg.start//60):02d}:{int(seg.start%60):02d}] {seg.text.strip()}"
        print(line)
        f.write(line + "\n")
