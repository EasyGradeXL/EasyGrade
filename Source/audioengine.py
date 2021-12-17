from __future__ import absolute_import
from __future__ import division

import pyaudio
from six.moves import queue

class AudioEngine(object):
    """
    This class generates chunks of audio to feed to the speech recognition engine

    Methods
    _______
        pause()
            Pauses the audio stream
        resume()
            Resumes the audio stream
    """

    def __init__(self, sampleRate: int, chunkSize: int):
        """
        Parameters
        ----------
        sampleRate : int
            Audio sample rate. Google recommends 16 kHz
        sound : str
            The audio queue chuck size in bytes
        """

        self.__sampleRate = sampleRate
        self.__chunkSize = chunkSize
        # Initialize the audio queue
        self.__openAudioStream()

    def __openAudioStream(self):
        """Opens the audio stream and initiates the chunk queue.
        """

        self.__chunkQueue = queue.Queue()                  # Queue that holds the audio chunks
        self.__audioInterface = pyaudio.PyAudio()          # Create audio interface
        self.audioGenerator = self.__chunkGenerator()      # Callback function to generate the audio chunks
        # Open audio stream
        self.__audioStream = self.__audioInterface.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=self.__sampleRate,
            input=True,
            frames_per_buffer=self.__chunkSize,
            stream_callback=self.__fillChunkBuffer,
        )
        self.__audioStream.start_stream()

    def __chunkGenerator(self):
        """ Chunk generator to fill audio queue

            Yields
            -------
            chunk : list
                A chunk of sampled audio data
        """

        while True:
            chunk = self.__chunkQueue.get(block=True)       # Get audio chunk from queue
            if chunk is None:
                return
            data = [chunk]
            # Now consume whatever other data's still in the queue
            while True:
                try:
                    chunk = self.__chunkQueue.get(block=False)
                    if chunk is None:
                        return
                    data.append(chunk)
                except queue.Empty:
                    break
            yield b"".join(data)

    def __fillChunkBuffer(self, in_data: list, frame_count, time_info, status_flags):
        """ Continuously collect data from the audio stream and adds it to self.__chunkQueue
        Parameters
        ----------
        in_data : list
            Audio chunks from the chunk generator

        Returns
        _______
        pyaudio.paContinue : Portaudio callback return code
            Indicates __chunkGenerator must continue sending chunks
        """
        self.__chunkQueue.put(in_data)
        return None, pyaudio.paContinue

    def pause(self):
        """ Temporarily stops the audio stream"""

        self.__audioStream.stop_stream()

    def resume(self):
        """ Resumes the audio stream"""

        self.__audioStream.start_stream()

    def __del__(self):
        """
        Stops and closes the audio stream and terminates the PyAudio interface
        """

        try:
            self.__audioStream.stop_stream()
        except:
            pass
        try:
            self.__audioStream.close()
        except:
            pass
        try:
            self.__chunkQueue.put(None)
        except:
            pass
        try:
            self.__audioInterface.terminate()
        except:
            pass
