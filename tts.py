# V3
import os
import torch
import time
from playsound import playsound
import sounddevice as sd

device = torch.device('cpu')
torch.set_num_threads(4)
local_file = 'model.pt'

if not os.path.isfile(local_file):
    print('downloading model')
    torch.hub.download_url_to_file('https://models.silero.ai/models/tts/ru/v4_ru.pt',
                                   local_file)

model = torch.package.PackageImporter(local_file).load_pickle("tts_models", "model")
model.to(device)
print('model loaded')


def play_sound(text):
    sample_rate = 48000
    speaker = 'xenia'
    put_accent = True
    put_yo = True
    audio = model.apply_tts(text=text,
                                 speaker=speaker,
                                 sample_rate=sample_rate,
                                 put_accent=True,
                                 put_yo=True)
    #playsound(audio_paths)
    sd.play(audio, sample_rate * 1.05)
    time.sleep((len(audio) / sample_rate))
    sd.stop()
def play_ssml_sound(text):
    sample_rate = 48000
    speaker = 'xenia'
    put_accent = True
    put_yo = True
    audio = model.apply_tts(ssml_text=text,
                                 speaker=speaker,
                                 sample_rate=sample_rate,
                                 put_accent=True,
                                 put_yo=True)
    #playsound(audio_paths)
    sd.play(audio, sample_rate * 1.05)
    time.sleep((len(audio) / sample_rate))
    sd.stop()