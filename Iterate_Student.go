package main

import (
	"context"
    "log"
	"time"
	"fmt"

    "github.com/chromedp/chromedp"
	"github.com/chromedp/cdproto/dom"
)

func main() {
	opts := append(chromedp.DefaultExecAllocatorOptions[:],
		chromedp.Flag("headless", false),
		chromedp.Flag("disable-gpu", false),
		chromedp.Flag("enable-automation", false),
		chromedp.Flag("disable-extensions", false),
	)
	
	allocCtx, cancel := chromedp.NewExecAllocator(context.Background(), opts...)
	defer cancel()
	
	// create context
	ctx, cancel := chromedp.NewContext(allocCtx, chromedp.WithLogf(log.Printf))
	defer cancel()

	var html string
	
	if err := chromedp.Run(ctx,
		chromedp.Navigate("https://scrapingclub.com/exercise/list_infinite_scroll"),
		// wait for the page to load
		chromedp.Sleep(2000*time.Millisecond),
		// extract the raw HTML from the page
		chromedp.ActionFunc(func(ctx context.Context) error {
			// select the root node on the page
			rootNode, err := dom.GetDocument().Do(ctx)
			if err != nil {
			   return err
			}
			
			html, err = dom.GetOuterHTML().WithNodeID(rootNode.NodeID).Do(ctx)
			return err
		 }),
	); err != nil {
		log.Fatal(err)
	}

	fmt.Println(html)
}