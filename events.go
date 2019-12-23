package exchange

// CalendarEvent models a calendar event
type calendarEvent struct {
	Title     string `json:"title"`
	StartTime string `json:"startTime"`
	EndTime   string `json:"endTime"`
}

// ExchangeToken models an exchange token
type token struct {
	Type       string `json:"token_type"`
	ExpireTime int    `json:"expires_in"`
	Token      string `json:"access_token"`
}

type eventResponse struct {
	Events []exchangeEvent `json:"value"`
}

// ExchangeEvent models an event returned by microsoft exchange service
type exchangeEvent struct {
	ID                         string       `json:"id"`
	CreatedDateTime            string       `json:"createdDateTime"`
	LastModifiedDateTime       string       `json:"lastModifiedDateTime"`
	ChangeKey                  string       `json:"changeKey"`
	Categories                 []string     `json:"categories"`
	OriginalStartTimeZone      string       `json:"originalStartTimeZone"`
	OriginalEndTimeZone        string       `json:"originalEndTimeZone"`
	ICalUID                    string       `json:"iCalUId"`
	ReminderMinutesBeforeStart int          `json:"reminderMinutesBeforeStart"`
	IsReminderOn               bool         `json:"isReminderOn"`
	HasAttachments             bool         `json:"hasAttachments"`
	Subject                    string       `json:"subject"`
	BodyPreview                string       `json:"bodyPreview"`
	Importance                 string       `json:"importance"`
	Sensitivity                string       `json:"sensitivity"`
	IsAllDay                   bool         `json:"isAllDay"`
	IsCancelled                bool         `json:"isCancelled"`
	IsOrganizer                bool         `json:"isOrganizer"`
	ResponseRequested          bool         `json:"responseRequested"`
	SeriesMasterID             string       `json:"seriesMasterId"`
	ShowAs                     string       `json:"showAs"`
	EventType                  string       `json:"type"`
	WebLink                    string       `json:"webLink"`
	OnlineMeetingURL           string       `json:"onlineMeetingUrl"`
	Recurrence                 string       `json:"recurrence"`
	Body                       eventBody    `json:"body"`
	Start                      exchangeDate `json:"start"`
	End                        exchangeDate `json:"end"`
}

type eventRequest struct {
	Subject   string       `json:"Subject"`
	Body      eventBody    `json:"Body"`
	Start     exchangeDate `json:"Start"`
	End       exchangeDate `json:"End"`
	Attendees []attendee   `json:"Attendees"`
}

type eventBody struct {
	ContentType string `json:"ContentType"`
	Content     string `json:"Content"`
}

type exchangeDate struct {
	DateTime string `json:"DateTime"`
	TimeZone string `json:"TimeZone"`
}

type attendee struct {
	EmailAddress emailAddress `json:"EmailAddress"`
	Type         string       `json:"Type"`
}

type emailAddress struct {
	Address string `json:"Address"`
	Name    string `json:"Name"`
}

type calendarResponse struct {
	Calendars []calendar `json:"value"`
}

type calendar struct {
	ID                            string        `json:"id"`
	Name                          string        `json:"name"`
	Color                         string        `json:"color"`
	IsDefault                     bool          `json:"isDefaultCalendar"`
	ChangeKey                     string        `json:"changeKey"`
	CanShare                      bool          `json:"canShare"`
	CanViewPrivate                bool          `json:"canViewPrivateItems"`
	CanEdit                       bool          `json:"canEdit"`
	AllowedOnlineMeetingProviders []string      `json:"allowedOnlineMeetingProviders"`
	DefaultOnlineMeetingProvider  string        `json:"defaultOnlineMeetingProvider"`
	TallyingResponses             bool          `json:"isTallyingResponses"`
	Removable                     bool          `json:"isRemovable"`
	Owner                         calendarOwner `json:"owner"`
}

type calendarOwner struct {
	Name    string `json:"name"`
	Address string `json:"Address"`
}
