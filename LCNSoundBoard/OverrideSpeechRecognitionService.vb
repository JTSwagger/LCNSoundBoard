Imports System.Threading
Imports Microsoft.ProjectOxford
Imports Microsoft.ProjectOxford.SpeechRecognition

Public Class OverrideSpeechRecognitionService
    Implements IDisposable

    Private m_FinalResponseEvent As AutoResetEvent
    Public m_recoMode As SpeechRecognitionMode
    Public m_isMicrophoneReco As Boolean
    Public m_micClient As MicrophoneRecognitionClient

    Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
        m_FinalResponseEvent.Dispose()
    End Sub

    Public Sub OnResponseReceivedHandler(sender As Object, ByRef e As SpeechResponseEventArgs)

        Dim isFinalDictationMessage As Boolean = m_recoMode = SpeechRecognitionMode.LongDictation And (
        e.PhraseResponse.RecognitionStatus = RecognitionStatus.EndOfDictation Or
        e.PhraseResponse.RecognitionStatus = RecognitionStatus.DictationEndSilenceTimeout
        )

        If m_isMicrophoneReco And ((m_recoMode = SpeechRecognitionMode.ShortPhrase) Or isFinalDictationMessage) Then
            m_micClient.EndMicAndRecognition()
        End If

        If isFinalDictationMessage = False Then
            Console.WriteLine("**** NBEST RESULTS ****")
            For i As Integer = 0 To e.PhraseResponse.Results.Length
                Console.WriteLine("[{0}] Confidence={1} Text={2}", i, e.PhraseResponse.Results(i).Confidence,
                e.PhraseResponse.Results(i).DisplayText)
            Next
            Console.WriteLine(" ")
        End If

    End Sub

    Public Sub OnIntentHandler(sender As Object, e As SpeechIntentEventArgs)
        Console.WriteLine("**** FINAL INTENT ****")
        Console.WriteLine(e.Payload.ToString)
        Console.WriteLine(" ")
    End Sub

    Public Sub OnPartialResponseReceivedHandler(sender As Object, e As PartialSpeechResponseEventArgs)
        Console.WriteLine("**** PARTIAL RESULT ****")
        Console.WriteLine(e.PartialResult.ToString)
        Console.WriteLine(" ")
    End Sub

    Public Sub OnConversationErrorHandler(sender As Object, e As SpeechErrorEventArgs)
        Console.WriteLine("**** ERROR DETECTED ****")
        Console.WriteLine(e.SpeechErrorCode.ToString)
        Console.WriteLine(e.SpeechErrorText)
        Console.WriteLine(" ")
    End Sub

    Public Sub OnMicrophoneStatus(sender As Object, e As MicrophoneEventArgs)
        Console.WriteLine("*** MIC STATUS: " + e.Recording + " ****")
        If e.Recording == False Then
            m_micClient.EndMicAndRecognition()
        End If
        Console.WriteLine(" ")
    End Sub

    Public Sub DoDataRecognition(dataClient As DataRecognitionClient)
        Console.WriteLine("Screw it, we're not actually using this method anyhow.")
    End Sub

    Public Function createMicClient(Optional ByVal useDefaultEventHandlers As Boolean = True) As MicrophoneRecognitionClient

    End Function

    ' Optional ByVal lets us set default parameter values. 200 and 15 are the API defaults,
    ' but we want to be able to set them at will. So, win/win.
    Sub DoMicrophoneRecognition(micClient As MicrophoneRecognitionClient,
    Optional ByVal shortWait As Integer = 200, Optional ByVal longWait As Integer = 15)
        Dim waitSeconds As Integer = 0

        If m_recoMode = SpeechRecognitionMode.LongDictation Then
            waitSeconds = longWait
        Else
            waitSeconds = shortWait
        End If

        Try
            micClient.StartMicAndRecognition()
            Console.WriteLine("Start talking... >")

            Dim isReceivedResponse As Boolean = m_FinalResponseEvent.WaitOne(waitSeconds * 1000)
            If isReceivedResponse = False Then
                Console.WriteLine("{0}: timed out waiting for conversation response after {1}ms".Format(
                    DateTime.UtcNow(), waitSeconds * 1000
                ))
            End If
        Finally
            micClient.EndMicAndRecognition()
        End Try
    End Sub

    Public Sub main()
        Dim recoEx As New OverrideSpeechRecognitionService()
        recoEx.m_recoMode = SpeechRecognitionMode.ShortPhrase
        recoEx.m_isMicrophoneReco = True
        ' WTF 
    End Sub
End Class
