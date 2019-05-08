https://stackoverflow.com/questions/8680640/vba-how-to-conditionally-skip-a-for-loop-iteration

<pre>

<b>For i</b> = 1 <b>To</b> n: <b><em>Do</em></b>

    <b>If</b> <em>condition</em> <b>Then</b> <b><em>Exit Do</em></b> <em>'Exit Do is the Continue</em>

    Debug.Print i

<b><em>Loop While False:</em></b> <b>Next i</b>
</pre>

