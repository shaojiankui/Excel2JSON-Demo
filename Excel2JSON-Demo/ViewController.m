//
//  ViewController.m
//  Excel2JSON-Demo
//
//  Created by Jakey on 2017/10/11.
//  Copyright © 2017年 Jakey. All rights reserved.
//

#import "ViewController.h"
#import "DHxlsReader.h"

@interface ViewController ()

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    // Do any additional setup after loading the view, typically from a nib.

    NSString *path = [[[NSBundle mainBundle] resourcePath] stringByAppendingPathComponent:@"shici.xls"];


    // xls_debug = 1; // good way to see everything in the Excel file

    DHxlsReader *reader = [DHxlsReader xlsReaderWithPath:path];
    assert(reader);

    NSString *text = @"";

    text = [text stringByAppendingFormat:@"AppName: %@\n", reader.appName];
    text = [text stringByAppendingFormat:@"Author: %@\n", reader.author];
    text = [text stringByAppendingFormat:@"Category: %@\n", reader.category];
    text = [text stringByAppendingFormat:@"Comment: %@\n", reader.comment];
    text = [text stringByAppendingFormat:@"Company: %@\n", reader.company];
    text = [text stringByAppendingFormat:@"Keywords: %@\n", reader.keywords];
    text = [text stringByAppendingFormat:@"LastAuthor: %@\n", reader.lastAuthor];
    text = [text stringByAppendingFormat:@"Manager: %@\n", reader.manager];
    text = [text stringByAppendingFormat:@"Subject: %@\n", reader.subject];
    text = [text stringByAppendingFormat:@"Title: %@\n", reader.title];

    text = [text stringByAppendingFormat:@"\n\nNumber of Sheets: %u\n", reader.numberOfSheets];

#if 1
    [reader startIterator:0];

    while(YES) {
        DHcell *cell = [reader nextCell];
        if(cell.type == cellBlank) break;

        text = [text stringByAppendingFormat:@"\n%@\n", [cell dump]];
    }
#else
    int row = 2;
    while(YES) {
        DHcell *cell = [reader cellInWorkSheetIndex:0 row:row col:2];
        if(cell.type == cellBlank) break;
        DHcell *cell1 = [reader cellInWorkSheetIndex:0 row:row col:3];
        NSLog(@"\nCell:%@\nCell1:%@\n", [cell dump], [cell1 dump]);
        row++;

        //text = [text stringByAppendingFormat:@"\n%@\n", [cell dump]];
        //text = [text stringByAppendingFormat:@"\n%@\n", [cell1 dump]];
    }
#endif
    NSLog(@"%@",text);
    NSMutableDictionary *excelDic = [NSMutableDictionary dictionary];
    for(int i=0;i<reader.numberOfSheets;i++){
        NSMutableArray *headerArray = [NSMutableArray array];
        NSString *sheetName =  [reader sheetNameAtIndex:i];
        NSMutableArray *items = [([excelDic objectForKey:sheetName]?:[NSMutableArray array]) mutableCopy];
//        [reader startIterator:i];
        
        for (int r=1; r<[reader numberOfRowsInSheet:i]; r++) {
            NSMutableDictionary *oneRowDic = [NSMutableDictionary dictionary];
            for (int c=1; c<[reader numberOfColsInSheet:i]; c++) {
                DHcell *cell = [reader cellInWorkSheetIndex:i row:r col:c];
                if (r==1) {
                    [headerArray addObject:cell.str?:@"N"];
                }else{
                    [oneRowDic setObject:cell.str?:@"N" forKey:[headerArray objectAtIndex:c-1]];
                    [items addObject:oneRowDic];
                }
            }
        }
        [excelDic setObject:items forKey:sheetName];
    }
    
   
    NSData *jsonData = [NSJSONSerialization dataWithJSONObject:excelDic
                                                       options:NSJSONWritingPrettyPrinted
                                                         error:NULL];

    NSString *jsonString = [[NSString alloc] initWithData:jsonData encoding:NSUTF8StringEncoding];
    NSLog(@"jsonString:%@",jsonString);
}


- (void)didReceiveMemoryWarning {
    [super didReceiveMemoryWarning];
    // Dispose of any resources that can be recreated.
}


@end
