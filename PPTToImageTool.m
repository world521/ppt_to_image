//
//  PPTToImageTool.m
//  PPTToImage
//
//  Created by alex on 2019/7/18.
//  Copyright © 2019 alex. All rights reserved.
//

#import "PPTToImageTool.h"
#import <WebKit/WebKit.h>

/** scrollView适配iOS11 */
#define  adjustsScrollViewInsets(scrollView)\
do {\
_Pragma("clang diagnostic push")\
_Pragma("clang diagnostic ignored \"-Warc-performSelector-leaks\"")\
if ([scrollView respondsToSelector:NSSelectorFromString(@"setContentInsetAdjustmentBehavior:")]) {\
    NSMethodSignature *signature = [UIScrollView instanceMethodSignatureForSelector:@selector(setContentInsetAdjustmentBehavior:)];\
    NSInvocation *invocation = [NSInvocation invocationWithMethodSignature:signature];\
    NSInteger argument = 2;\
    invocation.target = scrollView;\
    invocation.selector = @selector(setContentInsetAdjustmentBehavior:);\
    [invocation setArgument:&argument atIndex:2];\
    [invocation retainArguments];\
    [invocation invoke];\
}\
_Pragma("clang diagnostic pop")\
} while (0)

@interface PPTToImageTool ()<WKNavigationDelegate>
@property (nonatomic, strong) UIWindow *window;
@property (nonatomic, strong) WKWebView *webV;
@property (nonatomic, assign) CGFloat webOffsetY;
@property (nonatomic, strong) NSMutableArray *pageInfoArr;          // 存储ppt每一页尺寸信息
@property (nonatomic, strong) NSMutableArray *pptImageArr;          // 存储ppt生成的图片
@property (nonatomic, copy) void (^callback)(NSArray *images);     // 图片回调
@property (nonatomic, copy) void (^progress)(CGFloat value);        // 转换进度
@end

@implementation PPTToImageTool

- (NSMutableArray *)pageInfoArr{
    if (!_pageInfoArr) {
        _pageInfoArr = [NSMutableArray array];
    }
    return _pageInfoArr;
}

- (NSMutableArray *)pptImageArr {
    if (!_pptImageArr) {
        _pptImageArr = [NSMutableArray array];
    }
    return _pptImageArr;
}

- (UIWindow *)window {
    if (!_window) {
        _window = [[UIWindow alloc] initWithFrame:CGRectMake(0, 0, UIScreen.mainScreen.bounds.size.width, 0)];
        _window.backgroundColor = [UIColor greenColor];
        _window.windowLevel = UIWindowLevelNormal;
        _window.hidden = YES;
    }
    return _window;
}

- (void)pptToImageWithPPTFileUrl:(NSString *)pptFileUrl progress:(void (^)(CGFloat value)) progress  completion:(void (^)(NSArray * images))completion{
    self.webOffsetY = 0;
    [self.pageInfoArr removeAllObjects];
    [self.pptImageArr removeAllObjects];
    self.webV = nil;
    self.callback = completion;
    self.progress = progress;
    
    [self p_initWebView];
    [self p_loadPPT:pptFileUrl];
}

// 初始化,并注册JS调用方法
- (void)p_initWebView {
    WKUserContentController *uController = [[WKUserContentController alloc] init];
    WKWebViewConfiguration *config = [[WKWebViewConfiguration alloc] init];
    config.userContentController = uController;
    WKWebView *webView = [[WKWebView alloc] initWithFrame:self.window.bounds configuration:config];
    webView.backgroundColor = [UIColor whiteColor];
    adjustsScrollViewInsets(webView.scrollView);
    webView.navigationDelegate = self;
    self.webV = webView;
    [self.window addSubview:webView];
}

// 加载PPT
- (void)p_loadPPT:(NSString *)pptFileUrl {
    if ([[UIDevice currentDevice].systemVersion floatValue] >= 9.0) {
        NSURL *fileURL = [NSURL fileURLWithPath:pptFileUrl];
        [self.webV loadFileURL:fileURL allowingReadAccessToURL:fileURL];
    }
}

// 查看 html结构
- (void)p_printHtml{
    [self.webV evaluateJavaScript:@"document.documentElement.innerHTML" completionHandler:^(id _Nullable html, NSError * _Nullable error) {
        NSLog(@"%@",html);
    }];
}

// 获取 ppt 信息,PPT页数,以及尺寸信息
- (void)p_getTTPInfo {
    dispatch_semaphore_t semaphore = dispatch_semaphore_create(0);
        
    // 清除间距 获取PPT页数
    NSString *jsStr = @"var eles = document.getElementsByClassName('slide'); for (var i=0; i<eles.length; i++) { eles[i].style.marginTop = '0'; }; document.body.style.margin = '0'; document.body.style.padding = '0'; document.getElementsByClassName('slide').length";
    __block NSInteger pageCount = 0;
    [self.webV evaluateJavaScript:jsStr completionHandler:^(id _Nullable pageC, NSError * _Nullable error) {
        dispatch_semaphore_signal(semaphore);
        pageCount = [pageC integerValue];
    }];
    while (dispatch_semaphore_wait(semaphore, DISPATCH_TIME_NOW)) {
        [[NSRunLoop currentRunLoop] runMode:NSDefaultRunLoopMode beforeDate:[NSDate dateWithTimeIntervalSinceNow:0.2]];
    }
    
    if (pageCount == 0) return;
    
    // 获取每一页信息
    for (NSInteger i = 0; i< pageCount; i++ ){
        __block NSInteger pageWidth = 0;
        __block NSInteger pageHeight = 0;
        
        NSString *getSizeJS = [NSString stringWithFormat:@"var slideE = window.getComputedStyle(document.getElementsByClassName('slide')[%ld]); var slideS = {'width': parseInt(slideE.width), 'height': parseInt(slideE.height)}; slideS", i];
        [self.webV evaluateJavaScript:getSizeJS completionHandler:^(id _Nullable pageSizeDic, NSError * _Nullable error) {
            pageWidth = [pageSizeDic[@"width"] integerValue];
            pageHeight = [pageSizeDic[@"height"] integerValue];
            dispatch_semaphore_signal(semaphore);
        }];
        while (dispatch_semaphore_wait(semaphore, DISPATCH_TIME_NOW)) {
            [[NSRunLoop currentRunLoop] runMode:NSDefaultRunLoopMode beforeDate:[NSDate dateWithTimeIntervalSinceNow:0.1]];
        }
        [self.pageInfoArr addObject:@(CGSizeMake(pageWidth, pageHeight))];
    }
    
    [self p_cropImage:0 maxCount:pageCount];
}

// 通过绘制进行截图
- (void)p_cropImage:(NSInteger)curIndex maxCount:(NSInteger)maxCount {
    if (curIndex >= maxCount) {
        if (self.callback) {
            self.window = nil;
            self.webV = nil;
            self.callback(self.pptImageArr);
        }
    } else {
        CGSize size = [self.pageInfoArr[curIndex] CGSizeValue];
        CGFloat targetW = self.webV.frame.size.width;
        CGFloat targetH = targetW * (size.height * 1.0 / size.width);
        self.webV.frame = CGRectMake(0, 0, targetW, targetH);
        self.webV.scrollView.contentOffset = CGPointMake(0, self.webOffsetY);
        self.webOffsetY += targetH;
        
        dispatch_after(dispatch_time(DISPATCH_TIME_NOW, (int64_t)(0.1 * NSEC_PER_SEC)), dispatch_get_main_queue(), ^{
            UIImage *pptImage = [self imageWithView:self.webV frame:self.webV.bounds];
            [self.pptImageArr addObject:pptImage];
            
            if (self.progress) {
                self.progress((CGFloat)self.pptImageArr.count / self.pageInfoArr.count);
            }
            
            [self p_cropImage:curIndex + 1 maxCount:maxCount];
        });
    }
}

// 截取响应视图
- (UIImage* )imageWithView:(UIView *)view frame:(CGRect)frame {
    @autoreleasepool {
        UIGraphicsBeginImageContextWithOptions(frame.size, YES, 0);
        CGContextRef context = UIGraphicsGetCurrentContext();
        if (!context) return nil;
        [view.layer renderInContext:context];
        UIImage *image = UIGraphicsGetImageFromCurrentImageContext();
        UIGraphicsEndImageContext();
        NSData *imageData = UIImageJPEGRepresentation(image, 0.5);
        UIImage *resultImage = [UIImage imageWithData:imageData];
        return resultImage;
    }
}

#pragma -mark UIWebViewDelegate

- (void)webView:(WKWebView *)webView didFinishNavigation:(WKNavigation *)navigation {
    [self p_printHtml];
    [self p_getTTPInfo];
}

@end
