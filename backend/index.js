const {exec} = require('child_process')
const {PowerShell} = require('node-powershell')
var fs = require('fs')
var bodyParser = require('body-parser')
var express = require('express')
const { constants } = require('buffer')

const app = express()
app.use(express.json())
app.listen(
    7777,
    '0.0.0.0'
)

var config = {}
// need to leave first line of config file blank, first field is undefined
fs.readFileSync('../config.txt', 'utf-16le').toString().split(/\r?\n/).forEach(line => {
    if(line.length==0 || line.startsWith('#') || line.startsWith('//'))
        return
    arr = line.split('=')
    config[arr[0].toString()]=arr[1]
})

const ps = new PowerShell({
    executionPolicy: 'Bypass',
    noProfile: true
})



app.get('/', (req,res) => {
    console.log("test request recieved")
    res.status(200).json({
        "success":true
    })
})


app.post('/change_slide', async (req,res) => {
    console.log(req['body'])
    
    const count = req.body.count
    const option = req.body.option
    const command = PowerShell.command`& "../PPT scripts/Change-All-PPT-Slides.ps1" -count ${count} -pptOption ${option}`
    const output = await ps.invoke(command)
    const stdout = output.stdout.toString()
    console.log(stdout)
    const statusCode = Number(stdout.substring(0,stdout.indexOf(':')))
    const msg = stdout.substring(stdout.indexOf(':')+1)
    res.status(statusCode).json({
        "message":"MEssage"
    })

    // console.log(stdout.toString('utf16le'))
    // const dict = JSON.parse(stdout.toString('utf8').replace('\n',' '))
    // console.log("stdout:"+typeof(dict)+" "+dict)
    // console.log(typeof(output.stderr.toString()))

    // const responseStr = stdout[stdout.length-1]
    // const statusCode = Number(responseStr.substring(0,responseStr.indexOf(':')))
    // const msg = responseStr.substring(responseStr.indexOf(':')+1)
    // res.status(statusCode).json(
    //     {
    //         "message":msg,
    //         "lyricsLang1":stdout[0],
    //         "lyricsLang2":stdout[1]
    //     }
    // )


    // exec('& "../PPT scripts/Change-All-PPT-Slides.ps1" -count 1', {'shell':'powershell.exe'},(error,stdout,stderr) => {
    //     console.log(error)
    //     console.log(stdout)
    //     console.log(stderr)
    // })
})


app.get('/create_ppt', (req,res) => {
    fs.access(config['songLyricsPath'], fs.constants.F_OK, (err) => {
        if(err){
            console.log(err)
        } else {
            fs.access(config['hymnLyricsPath'], fs.constants.F_OK, (error) => {
                if(error){
                    console.log(error)
                } else {
                    res.status(200).json({
                        "message":"PPT creation setup is ready"
                    })
                }
            })
        }
    })
})

app.get('/lyrics_files', (req,res) => {
    console.log("request lyrics")
    fs.readdir(config["songLyricsPath"], (err, songs) => {
        if(err){
            console.log(err)
        } else {
            fs.readdir(config["hymnLyricsPath"], (error, hymns) => {
                if(error){
                    console.log(error)
                } else {
                    statusCode = (songs.length + hymns.length == 0)? 204:200
                    console.log(songs)
                    console.log(hymns)
                    res.status(statusCode).json({
                        "songs" : songs,
                        "hymns" : hymns
                    })
                }
            })
        }
    })
})

app.get('/songs',(req,res) => {
    console.log("song request recieved")
    fs.readdir(config["songLyricsPath"], (err, files) => {
        if(err){
            console.log(err)
        } else {
            statusCode = (files.length == 0)? 204:200
            res.status(statusCode).json({
                songs : files
            })
        }
    })
})


app.get('/hymns', (req,res) => {
    console.log("hymn request recieved")
    fs.readdir(config["hymnLyricsPath"], (err, files) => {
        if(err){
            console.log(err)
        }else {
            statusCode = (files.length == 0)? 204:200
            res.status(statusCode).json({
                hymns : files
            })
        }
    })
})

app.post('/create_ppt',async (req,res) => {
    console.log("test ppt creation")
    isHymn = req.body.isHymn
    songList = req.body.songList
    startSlideShow = req.body.startSlideShow

    const command = PowerShell.command`& "../PPT Scripts/Create-PPT.ps1" -hymn ${isHymn} -songList ${songList} -startSlideShow ${startSlideShow}`
    const output = await ps.invoke(command)
    const stdout = output.stdout.toString().split('\n')
    console.log(stdout)
    const statusCode = Number(stdout[stdout.length-1].substring(0,stdout.indexOf(':')))
    const msg = stdout.substring(stdout[stdout.length-1].indexOf(':')+1)
    res.status(statusCode).json(
        {
            "message":msg
        }
    )
})