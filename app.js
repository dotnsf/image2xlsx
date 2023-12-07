//. app.js
var express = require( 'express' ),
    bodyParser = require( 'body-parser' ),
    fs = require( 'fs' ),
    multer = require( 'multer' ),
    { createCanvas, Image } = require( '@napi-rs/canvas' ),
    XLSX = require( 'xlsx-js-style' ),
    app = express();

app.use( express.static( __dirname + '/public/_doc' ) );
app.use( bodyParser.urlencoded( { extended: true } ) );
app.use( bodyParser.json() );
app.use( express.Router() );
app.use( multer( { dest: './tmp/' } ).single( 'file' ) );

//. CORS
var settings_cors = 'CORS' in process.env ? process.env.CORS : '';  //. "http://localhost:8080,https://xxx.herokuapp.com"
app.all( '/*', function( req, res, next ){
  if( settings_cors ){
    var origin = req.headers.origin;
    if( origin ){
      var cors = settings_cors.split( " " ).join( "" ).split( "," );

      //. cors = [ "*" ] への対応が必要
      if( cors.indexOf( '*' ) > -1 ){
        res.setHeader( 'Access-Control-Allow-Origin', '*' );
        res.setHeader( 'Access-Control-Allow-Methods', '*' );
        res.setHeader( 'Access-Control-Allow-Headers', '*' );
        res.setHeader( 'Vary', 'Origin' );
      }else{
        if( cors.indexOf( origin ) > -1 ){
          res.setHeader( 'Access-Control-Allow-Origin', origin );
          res.setHeader( 'Access-Control-Allow-Methods', '*' );
          res.setHeader( 'Access-Control-Allow-Headers', '*' );
          res.setHeader( 'Vary', 'Origin' );
        }
      }
    }
  }
  next();
});

app.get( '/ping', function( req, res ){
  res.contentType( 'application/json; charset=utf-8' );

  res.write( JSON.stringify( { status: true, message: 'PONG' }, null, 2 ) );
  res.end();
});

app.post( '/image', function( req, res ){
  var imgpath = req.file.path;
  var outputfilename = '';
  try{
    var imgtype = req.file.mimetype;
    var filename = req.file.originalname;
    //var imgsize = req.file.size;
    var ext = imgtype.split( "/" )[1];

    //. 出力ファイル名
    var tmp = filename.split( '.' );
    if( tmp.length > 1 ){ tmp.pop(); } //. 画像拡張子を取り除く
    outputfilename = tmp.join( '.' ) + '.xlsx';

    //. 画像解析
    var data = fs.readFileSync( imgpath );
    var img = new Image;
    img.src = data;

    var width = req.body.width ? parseInt( req.body.width ) : -1;
    var height = req.body.height ? parseInt( req.body.height ) : -1;
    if( width < 0 && height < 0 ){
      width = 100;  //. サイズの指定が何もない場合、幅100とする
      height = width * img.height / img.width;
    }else if( width < 0 ){
      width = height * img.width / img.height;
    }else if( height < 0 ){
      height = width * img.height / img.width;
    }

    //. リサイズして仮想キャンバスに描画
    var canvas = createCanvas( width, height );
    var ctx = canvas.getContext( '2d' );
    ctx.drawImage( img, 0, 0, img.width, img.height, 0, 0, width, height );

    //. 描画した画像のデータを取得
    var imagedata = ctx.getImageData( 0, 0, width, height );

    //. ワークブックを作成
    var wb = XLSX.utils.book_new();

    //. ワークシートを作成
    var row = [[]];
    var ws = XLSX.utils.aoa_to_sheet([row]);

    //. 必要なセルの行数と列数ぶんだけ、幅と高さを指定する
    var px = 2;  //. ２ピクセルずつ
    var wscols = [];
    for( var x = 0; x < imagedata.width; x ++ ){
      wscols.push( { wpx: px } );
    }
    var wsrows = [];
    for( var y = 0; y < imagedata.height; y ++ ){
      wsrows.push( { hpx: px } );
    }

    ws['!cols'] = wscols;
    ws['!rows'] = wsrows;

    //. １ピクセルずつ取り出して色を調べる
    for( var y = 0; y < imagedata.height; y ++ ){
      for( var x = 0; x < imagedata.width; x ++ ){
        var index = ( y * imagedata.width + x ) * 4;
        var rr = imagedata.data[index];    //. 0 <= xx < 256
        var gg = imagedata.data[index+1];
        var bb = imagedata.data[index+2];
        var alpha = imagedata.data[index+3];

        var rrggbb = rgb2hex( rr, gg, bb );
        var cname = cellname( x, y );
        ws[cname] = { v: "", f: undefined, t: 's', s: { fill: { fgColor: { rgb: rrggbb } } } };
      }
    }

    //. 値が適用される範囲を指定
    var cname = cellname( imagedata.width - 1, imagedata.height - 1 );
    ws['!ref'] = 'A1:' + cname;

    //. ワークシートをワークブックに追加
    XLSX.utils.book_append_sheet( wb, ws, filename );

    //. xlsx化
    XLSX.writeFileSync( wb, 'tmp/' + outputfilename );
    //res.redirect( './' + outputfilename );

    res.contentType( 'application/vnd.ms-excel' );
    res.setHeader( 'Content-Disposition', 'attachment; filename="' + outputfilename + '"' );
    var bin = fs.readFileSync( 'tmp/' + outputfilename );
    res.write( bin );
    res.end();

    //res.contentType( 'application/json; charset=utf-8' );
    //res.write( JSON.stringify( { status: true }, null, 2 ) );
    //res.end();
  }catch( e ){
    console.log( e );
    res.contentType( 'application/json; charset=utf-8' );
    res.status( 400 );
    res.write( JSON.stringify( { status: false, error: e } ) );
    res.end();
  }finally{
    fs.unlinkSync( imgpath );
    fs.unlinkSync( 'tmp/' + outputfilename );
  }
});

function color2hex( c ){
  var x = c.toString( 16 );
  if( x.length == 1 ){ x = '0' + x; }
  return x;
}

function rgb2hex( r, g, b ){
  return color2hex( r ) + color2hex( g ) + color2hex( b );
}

function cellname( x, y ){
  //. String.fromCharCode( 65 ) = 'A';
  var c = '';
  var r = ( y + 1 );

  while( x >= 0 ){
    var m = x % 26;
    x = Math.floor( x / 26 );

    c = String.fromCharCode( 65 + m ) + c;
    if( x == 0 ){
      x = -1;
    }else{
      x --;
    }
  }

  return ( c + r );
}


var port = process.env.PORT || 8080;
app.listen( port );
console.log( "server starting on " + port + " ..." );

module.exports = app;
